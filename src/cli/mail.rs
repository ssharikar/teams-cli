use clap::{Args, Subcommand};
use std::io::Read;
use std::time::Instant;

use crate::api::outlook::OutlookClient;
use crate::api::HttpClient;
use crate::auth::token::TokenSet;
use crate::error::{Result, TeamsError};
use crate::html::strip_html;
use crate::models::outlook::{
    DraftMessage, FileAttachment, ItemBody, Recipient, SendMailMessage, SendMailRequest,
};
use crate::output::{self, OutputFormat};

#[derive(Args)]
pub struct MailArgs {
    #[command(subcommand)]
    pub command: MailCommand,
}

#[derive(Subcommand)]
pub enum MailCommand {
    /// List inbox messages
    List {
        /// Mail folder to list (default: Inbox)
        #[arg(long, default_value = "Inbox")]
        folder: String,
        /// Show messages from the past duration (e.g. "1h", "24h", "7d")
        #[arg(long)]
        since: Option<String>,
        /// Only show unread messages
        #[arg(long)]
        unread: bool,
        /// Maximum number of messages to fetch
        #[arg(long, default_value = "25")]
        limit: u32,
    },
    /// Read a specific email
    Read {
        /// Message ID
        message_id: String,
    },
    /// Permanently delete a message (e.g. a draft)
    Delete {
        /// Message ID
        message_id: String,
    },
    /// Send an existing draft (e.g. one created by `mail draft`)
    SendDraft {
        /// Draft message ID
        message_id: String,
    },
    /// Send an email
    Send {
        /// Recipient email addresses (repeatable, or comma-separated)
        #[arg(long, required = true, value_delimiter = ',')]
        to: Vec<String>,
        /// CC recipients (repeatable, or comma-separated)
        #[arg(long, value_delimiter = ',')]
        cc: Vec<String>,
        /// BCC recipients (repeatable, or comma-separated)
        #[arg(long, value_delimiter = ',')]
        bcc: Vec<String>,
        /// Subject line
        #[arg(long, required = true)]
        subject: String,
        /// Message body text
        #[arg(long)]
        body: Option<String>,
        /// Read body from stdin
        #[arg(long)]
        stdin: bool,
        /// Read body from file
        #[arg(long)]
        file: Option<String>,
        /// Send as HTML
        #[arg(long)]
        html: bool,
        /// Attach a file (repeatable)
        #[arg(short = 'a', long = "attach")]
        attach: Vec<String>,
    },
    /// Save an email to Drafts without sending it
    Draft {
        /// Recipient email addresses (repeatable, or comma-separated)
        #[arg(long, value_delimiter = ',')]
        to: Vec<String>,
        /// CC recipients (repeatable, or comma-separated)
        #[arg(long, value_delimiter = ',')]
        cc: Vec<String>,
        /// BCC recipients (repeatable, or comma-separated) — handles large
        /// lists (tested with ~180 recipients in one draft)
        #[arg(long, value_delimiter = ',')]
        bcc: Vec<String>,
        /// Subject line
        #[arg(long, required = true)]
        subject: String,
        /// Message body text
        #[arg(long)]
        body: Option<String>,
        /// Read body from stdin
        #[arg(long)]
        stdin: bool,
        /// Read body from file
        #[arg(long)]
        file: Option<String>,
        /// Save as HTML
        #[arg(long)]
        html: bool,
        /// Attach a file (repeatable)
        #[arg(short = 'a', long = "attach")]
        attach: Vec<String>,
    },
    /// Search emails
    Search {
        /// Search query (Outlook search syntax)
        query: String,
        /// Maximum results
        #[arg(long, default_value = "25")]
        limit: u32,
    },
}

pub async fn handle(
    args: &MailArgs,
    tokens: &TokenSet,
    http: &HttpClient,
    format: OutputFormat,
) -> Result<()> {
    let bearer = tokens.outlook_bearer()?;
    let client = OutlookClient::new(http, bearer);

    match &args.command {
        MailCommand::List {
            folder,
            since,
            unread,
            limit,
        } => {
            let mut filters: Vec<String> = Vec::new();

            if let Some(since_str) = since {
                let dt = parse_since(since_str)?;
                filters.push(format!("ReceivedDateTime ge {}", dt.to_rfc3339()));
            }

            if *unread {
                filters.push("IsRead eq false".to_string());
            }

            let filter = if filters.is_empty() {
                None
            } else {
                Some(filters.join(" and "))
            };

            let start = Instant::now();
            let messages = client
                .list_messages(folder, filter.as_deref(), *limit)
                .await?;

            let display: Vec<serde_json::Value> = messages
                .iter()
                .map(|m| {
                    let from = m.from.as_ref().map(|r| r.display()).unwrap_or_default();
                    serde_json::json!({
                        "id": m.id,
                        "from": from,
                        "subject": m.subject,
                        "date": m.received_date_time,
                        "preview": m.body_preview,
                        "read": m.is_read,
                        "attachments": m.has_attachments,
                    })
                })
                .collect();

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        MailCommand::Read { message_id } => {
            let start = Instant::now();
            let msg = client.get_message(message_id).await?;

            let body_text = msg
                .body
                .as_ref()
                .map(|b| {
                    if b.content_type == "HTML" {
                        strip_html(&b.content)
                    } else {
                        b.content.clone()
                    }
                })
                .unwrap_or_default();

            let from = msg.from.as_ref().map(|r| r.display()).unwrap_or_default();
            let to: Vec<String> = msg.to_recipients.iter().map(|r| r.display()).collect();
            let cc: Vec<String> = msg.cc_recipients.iter().map(|r| r.display()).collect();
            let bcc: Vec<String> = msg.bcc_recipients.iter().map(|r| r.display()).collect();

            let display = serde_json::json!({
                "id": msg.id,
                "subject": msg.subject,
                "from": from,
                "to": to,
                "cc": cc,
                "bcc": bcc,
                "date": msg.received_date_time,
                "body": body_text,
                "read": msg.is_read,
                "attachments": msg.has_attachments,
            });

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        MailCommand::Delete { message_id } => {
            let start = Instant::now();
            client.delete_message(message_id).await?;

            let result = serde_json::json!({
                "status": "deleted",
                "id": message_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }

        MailCommand::SendDraft { message_id } => {
            let start = Instant::now();
            client.send_draft(message_id).await?;

            let result = serde_json::json!({
                "status": "sent",
                "id": message_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }

        MailCommand::Send {
            to,
            cc,
            bcc,
            subject,
            body,
            stdin,
            file,
            html,
            attach,
        } => {
            let content = resolve_body(body.as_deref(), *stdin, file.as_deref())?;

            let content_type = if *html { "HTML" } else { "Text" };

            let request = SendMailRequest {
                message: SendMailMessage {
                    subject: subject.clone(),
                    body: ItemBody {
                        content_type: content_type.to_string(),
                        content,
                    },
                    to_recipients: recipients_from(to),
                    cc_recipients: recipients_from(cc),
                    bcc_recipients: recipients_from(bcc),
                    attachments: attachments_from(attach)?,
                },
                save_to_sent_items: true,
            };

            let start = Instant::now();
            client.send_message(&request).await?;

            let result = serde_json::json!({
                "status": "sent",
                "to": to,
                "cc": cc,
                "bcc": bcc,
                "subject": subject,
                "attachments": attach,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }

        MailCommand::Draft {
            to,
            cc,
            bcc,
            subject,
            body,
            stdin,
            file,
            html,
            attach,
        } => {
            if to.is_empty() && cc.is_empty() && bcc.is_empty() {
                return Err(TeamsError::InvalidInput(
                    "provide at least one of --to, --cc, --bcc".into(),
                ));
            }

            let content = resolve_body(body.as_deref(), *stdin, file.as_deref())?;

            let content_type = if *html { "HTML" } else { "Text" };

            let draft = DraftMessage {
                subject: subject.clone(),
                body: ItemBody {
                    content_type: content_type.to_string(),
                    content,
                },
                to_recipients: recipients_from(to),
                cc_recipients: recipients_from(cc),
                bcc_recipients: recipients_from(bcc),
                attachments: attachments_from(attach)?,
            };

            let start = Instant::now();
            let created = client.create_draft(&draft).await?;

            let result = serde_json::json!({
                "status": "draft",
                "id": created.id,
                "to": to,
                "cc": cc,
                "bcc": bcc,
                "subject": subject,
                "attachments": attach,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }

        MailCommand::Search { query, limit } => {
            let start = Instant::now();
            let messages = client.search_messages(query, *limit).await?;

            let display: Vec<serde_json::Value> = messages
                .iter()
                .map(|m| {
                    let from = m.from.as_ref().map(|r| r.display()).unwrap_or_default();
                    serde_json::json!({
                        "id": m.id,
                        "from": from,
                        "subject": m.subject,
                        "date": m.received_date_time,
                        "preview": m.body_preview,
                        "read": m.is_read,
                    })
                })
                .collect();

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }
    }

    Ok(())
}

fn recipients_from(addrs: &[String]) -> Vec<Recipient> {
    addrs
        .iter()
        .map(|addr| Recipient::new(addr.trim()))
        .collect()
}

/// Reads each path and base64-encodes it for Outlook's inline attachment
/// form. Fine for typical attachment sizes; a very large file would need
/// the separate chunked-upload session API, which isn't implemented here.
fn attachments_from(paths: &[String]) -> Result<Vec<FileAttachment>> {
    use base64::Engine;

    paths
        .iter()
        .map(|path| {
            let bytes = std::fs::read(path)
                .map_err(|e| TeamsError::InvalidInput(format!("cannot read file {path:?}: {e}")))?;
            let name = std::path::Path::new(path)
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or(path)
                .to_string();
            let content_bytes = base64::engine::general_purpose::STANDARD.encode(bytes);
            Ok(FileAttachment::new(name, content_bytes))
        })
        .collect()
}

fn resolve_body(body: Option<&str>, stdin: bool, file: Option<&str>) -> Result<String> {
    if stdin {
        let mut buf = String::new();
        std::io::stdin()
            .read_to_string(&mut buf)
            .map_err(|e| TeamsError::InvalidInput(format!("stdin: {e}")))?;
        return Ok(buf.trim_end().to_string());
    }

    if let Some(path) = file {
        return std::fs::read_to_string(path)
            .map(|s| s.trim_end().to_string())
            .map_err(|e| TeamsError::InvalidInput(format!("file {path}: {e}")));
    }

    if let Some(text) = body {
        return Ok(text.to_string());
    }

    Err(TeamsError::InvalidInput(
        "provide --body, --stdin, or --file".into(),
    ))
}

fn parse_since(since: &str) -> Result<chrono::DateTime<chrono::Utc>> {
    let s = since.trim();
    if s.len() < 2 {
        return Err(TeamsError::InvalidInput(format!(
            "invalid duration '{s}': expected format like 1h, 24h, 7d"
        )));
    }

    let (num_str, unit) = s.split_at(s.len() - 1);
    let num: i64 = num_str.parse().map_err(|_| {
        TeamsError::InvalidInput(format!(
            "invalid duration '{s}': expected format like 1h, 24h, 7d"
        ))
    })?;

    let duration = match unit {
        "m" => chrono::Duration::minutes(num),
        "h" => chrono::Duration::hours(num),
        "d" => chrono::Duration::days(num),
        _ => {
            return Err(TeamsError::InvalidInput(format!(
                "invalid duration unit '{unit}': expected m, h, or d"
            )))
        }
    };

    Ok(chrono::Utc::now() - duration)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_since_hours() {
        let dt = parse_since("1h").unwrap();
        let diff = chrono::Utc::now() - dt;
        assert!(diff.num_minutes() >= 59 && diff.num_minutes() <= 61);
    }

    #[test]
    fn parse_since_days() {
        let dt = parse_since("7d").unwrap();
        let diff = chrono::Utc::now() - dt;
        assert!(diff.num_days() >= 6 && diff.num_days() <= 7);
    }

    #[test]
    fn parse_since_minutes() {
        let dt = parse_since("30m").unwrap();
        let diff = chrono::Utc::now() - dt;
        assert!(diff.num_minutes() >= 29 && diff.num_minutes() <= 31);
    }

    #[test]
    fn parse_since_invalid_unit() {
        assert!(parse_since("5x").is_err());
    }

    #[test]
    fn parse_since_invalid_number() {
        assert!(parse_since("abch").is_err());
    }

    #[test]
    fn parse_since_too_short() {
        assert!(parse_since("h").is_err());
    }

    #[test]
    fn resolve_body_from_body_flag() {
        let result = resolve_body(Some("hello"), false, None).unwrap();
        assert_eq!(result, "hello");
    }

    #[test]
    fn resolve_body_missing_all() {
        assert!(resolve_body(None, false, None).is_err());
    }
}
