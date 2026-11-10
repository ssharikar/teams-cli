use clap::{Args, Subcommand};
use std::io::Read;
use std::time::Instant;

use crate::api::blob::BlobClient;
use crate::api::messages::MessagesClient;
use crate::api::HttpClient;
use crate::auth::token::TokenSet;
use crate::error::Result;
use crate::html::strip_html;
use crate::output::{self, OutputFormat};

#[derive(Args)]
pub struct MessageArgs {
    #[command(subcommand)]
    pub command: MessageCommand,
}

#[derive(Subcommand)]
pub enum MessageCommand {
    /// List messages in a conversation
    List {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Maximum number of messages to fetch
        #[arg(long, default_value = "50")]
        limit: u32,
    },
    /// Send a message to a conversation
    Send {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Message body text
        #[arg(long)]
        body: Option<String>,
        /// Read message body from stdin
        #[arg(long)]
        stdin: bool,
        /// Send as HTML (auto-detected if content contains <at> mention tags)
        #[arg(long)]
        html: bool,
        /// Attach an image file to the message
        #[arg(long)]
        file: Option<String>,
    },
    /// Get a specific message
    Get {
        /// Conversation ID
        conversation_id: String,
        /// Message ID
        message_id: String,
    },
    /// Edit a message you previously sent
    Edit {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Message ID to edit
        message_id: String,
        /// New message body text
        #[arg(long)]
        body: Option<String>,
        /// Read new body from stdin
        #[arg(long)]
        stdin: bool,
        /// New body is HTML (auto-detected if content contains <at> mention tags)
        #[arg(long)]
        html: bool,
    },
    /// Delete a message you previously sent
    Delete {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Message ID to delete
        message_id: String,
    },
    /// React to a message
    React {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Message ID to react to
        message_id: String,
        /// Reaction type: like, heart, laugh, surprised, sad, angry
        #[arg(long)]
        reaction: String,
    },
    /// Remove a reaction from a message
    Unreact {
        /// Conversation ID (channel or chat thread ID)
        conversation_id: String,
        /// Message ID to remove reaction from
        message_id: String,
        /// Reaction type to remove: like, heart, laugh, surprised, sad, angry
        #[arg(long)]
        reaction: String,
    },
}

pub struct MessageContext<'a> {
    pub tokens: &'a TokenSet,
    pub messaging_token: &'a str,
    pub http: &'a HttpClient,
    pub chat_service_url: &'a str,
    pub ams_v2_url: &'a str,
    pub ams_url: &'a str,
}

pub async fn handle(
    args: &MessageArgs,
    ctx: &MessageContext<'_>,
    format: OutputFormat,
) -> Result<()> {
    let msg_client = MessagesClient::new(ctx.http, ctx.messaging_token, ctx.chat_service_url);

    match &args.command {
        MessageCommand::List {
            conversation_id,
            limit,
        } => {
            let start = Instant::now();
            let messages = msg_client.get_messages(conversation_id, *limit).await?;
            let display: Vec<serde_json::Value> = messages
                .iter()
                .filter(|m| m.message_type == "RichText/Html" || m.message_type == "Text")
                .map(|m| {
                    serde_json::json!({
                        "id": m.id,
                        "from": m.im_display_name,
                        "content": strip_html(&m.content),
                        "time": m.compose_time,
                        "type": m.message_type,
                    })
                })
                .collect();
            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }
        MessageCommand::Send {
            conversation_id,
            body,
            stdin,
            html,
            file,
        } => {
            let content = if *stdin {
                let mut buf = String::new();
                std::io::stdin()
                    .read_to_string(&mut buf)
                    .map_err(|e| crate::error::TeamsError::InvalidInput(format!("stdin: {e}")))?;
                Some(buf.trim_end().to_string())
            } else {
                body.clone()
            };

            // Upload image if --file is provided
            let mut amsreferences = None;
            let mut image_html = String::new();
            if let Some(file_path) = file {
                let path = std::path::Path::new(file_path);
                if !path.exists() {
                    return Err(crate::error::TeamsError::InvalidInput(format!(
                        "file not found: {file_path}"
                    )));
                }
                let blob_client =
                    BlobClient::new(ctx.http, ctx.messaging_token, ctx.ams_v2_url, ctx.ams_url);
                let blob_id = blob_client.upload_image(conversation_id, path).await?;
                image_html = blob_client.build_image_html(&blob_id);
                amsreferences = Some(vec![blob_id]);
            }

            // Build final content
            let final_raw = match (&content, image_html.is_empty()) {
                (Some(text), true) => text.clone(),
                (Some(text), false) => format!("{text}{image_html}"),
                (None, false) => image_html.clone(),
                (None, true) => {
                    return Err(crate::error::TeamsError::InvalidInput(
                        "provide --body, --stdin, or --file".into(),
                    ));
                }
            };

            // Auto-detect HTML if content contains <at> mention tags, --html flag, or image
            let is_html = *html || final_raw.contains("<at ") || amsreferences.is_some();
            let (final_content, mentions_json) = if is_html {
                parse_and_rewrite_mentions(&final_raw)
            } else {
                (final_raw, None)
            };

            let display_name = crate::auth::token::extract_username(&ctx.tokens.teams.raw)
                .unwrap_or_else(|_| "Unknown".into());

            let start = Instant::now();
            let result = msg_client
                .send_message(
                    conversation_id,
                    &final_content,
                    &display_name,
                    is_html,
                    mentions_json.as_deref(),
                    amsreferences,
                )
                .await?;
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        MessageCommand::Get {
            conversation_id,
            message_id,
        } => {
            let start = Instant::now();
            let messages = msg_client.get_messages(conversation_id, 200).await?;
            let message = messages
                .iter()
                .find(|m| m.id == *message_id)
                .ok_or_else(|| {
                    crate::error::TeamsError::NotFound(format!("message {message_id}"))
                })?;
            output::print_output(format, message, start.elapsed().as_millis() as u64);
        }
        MessageCommand::Edit {
            conversation_id,
            message_id,
            body,
            stdin,
            html,
        } => {
            let content = if *stdin {
                let mut buf = String::new();
                std::io::stdin()
                    .read_to_string(&mut buf)
                    .map_err(|e| crate::error::TeamsError::InvalidInput(format!("stdin: {e}")))?;
                buf.trim_end().to_string()
            } else {
                body.clone().ok_or_else(|| {
                    crate::error::TeamsError::InvalidInput("provide --body or --stdin".into())
                })?
            };

            let is_html = *html || content.contains("<at ");
            let (final_content, mentions_json) = if is_html {
                parse_and_rewrite_mentions(&content)
            } else {
                (content, None)
            };

            let display_name = crate::auth::token::extract_username(&ctx.tokens.teams.raw)
                .unwrap_or_else(|_| "Unknown".into());

            let start = Instant::now();
            msg_client
                .edit_message(
                    conversation_id,
                    message_id,
                    &final_content,
                    &display_name,
                    is_html,
                    mentions_json.as_deref(),
                )
                .await?;

            let result = serde_json::json!({
                "status": "edited",
                "conversation_id": conversation_id,
                "message_id": message_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        MessageCommand::Delete {
            conversation_id,
            message_id,
        } => {
            let start = Instant::now();
            msg_client
                .delete_message(conversation_id, message_id)
                .await?;

            let result = serde_json::json!({
                "status": "deleted",
                "conversation_id": conversation_id,
                "message_id": message_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        MessageCommand::React {
            conversation_id,
            message_id,
            reaction,
        } => {
            validate_reaction(reaction)?;
            let start = Instant::now();
            msg_client
                .react(conversation_id, message_id, reaction)
                .await?;
            let result = serde_json::json!({
                "message_id": message_id,
                "reaction": reaction,
                "action": "added",
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
        MessageCommand::Unreact {
            conversation_id,
            message_id,
            reaction,
        } => {
            validate_reaction(reaction)?;
            let start = Instant::now();
            msg_client
                .unreact(conversation_id, message_id, reaction)
                .await?;
            let result = serde_json::json!({
                "message_id": message_id,
                "reaction": reaction,
                "action": "removed",
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
    }
    Ok(())
}

const VALID_REACTIONS: &[&str] = &["like", "heart", "laugh", "surprised", "sad", "angry"];

fn validate_reaction(reaction: &str) -> Result<()> {
    if VALID_REACTIONS.contains(&reaction) {
        Ok(())
    } else {
        Err(crate::error::TeamsError::InvalidInput(format!(
            "invalid reaction '{}'. Valid reactions: {}",
            reaction,
            VALID_REACTIONS.join(", ")
        )))
    }
}

/// Minimal HTML entity unescape for display-name text.
fn html_unescape(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
}

/// Parse `<at id="8:orgid:...">Name</at>` tags from content.
/// Rewrites them to `<span>` mention tags and builds the mentions metadata JSON
/// that the Teams API expects. Only MRI-formatted IDs (8: for users, 28: for
/// bots) are recognized; non-MRI `<at>` tags are left untouched.
fn parse_and_rewrite_mentions(content: &str) -> (String, Option<String>) {
    let re_str = r#"<at id="([^"]+)">([^<]+)</at>"#;
    let re = regex::Regex::new(re_str).expect("valid regex");

    let mut mentions = Vec::new();

    // First pass: collect mentions with valid MRI prefixes
    for (idx, cap) in re.captures_iter(content).enumerate() {
        let mri = cap[1].to_string();
        // Only accept user MRIs (8:) and bot MRIs (28:)
        if mri.starts_with("8:") || mri.starts_with("28:") {
            let display_name = html_unescape(&cap[2]);
            mentions.push((mri, display_name, idx as u32));
        }
    }

    if mentions.is_empty() {
        return (content.to_string(), None);
    }

    // Second pass: rewrite <at> tags to <span> mention tags
    let mut rewritten = content.to_string();
    for (mri, name, id) in &mentions {
        let old_tag = format!(r#"<at id="{mri}">{name}</at>"#);
        // Also handle the unescaped form in case the original had entities
        let old_tag_escaped = format!(
            r#"<at id="{mri}">{}</at>"#,
            name.replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;")
        );
        let new_tag = format!(
            r#"<span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="{id}">{name}</span>"#
        );
        if rewritten.contains(&old_tag) {
            rewritten = rewritten.replacen(&old_tag, &new_tag, 1);
        } else {
            rewritten = rewritten.replacen(&old_tag_escaped, &new_tag, 1);
        }
    }

    // Wrap in <p> if not already wrapped — matches real Teams client output
    let trimmed = rewritten.trim();
    if !trimmed.starts_with("<p>") {
        rewritten = format!("<p>{trimmed}</p>");
    }

    // Build mentions metadata array
    let mentions_arr: Vec<serde_json::Value> = mentions
        .iter()
        .map(|(mri, name, id)| {
            serde_json::json!({
                "@type": "http://schema.skype.com/Mention",
                "itemid": id,
                "mri": mri,
                "mentionType": "person",
                "displayName": name,
            })
        })
        .collect();

    let mentions_json = serde_json::to_string(&mentions_arr).expect("serialize mentions");
    (rewritten, Some(mentions_json))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_mentions_no_at_tags() {
        let (content, mentions) = parse_and_rewrite_mentions("Hello world");
        assert_eq!(content, "Hello world");
        assert!(mentions.is_none());
    }

    #[test]
    fn parse_mentions_single_mention() {
        let input = r#"<at id="8:orgid:abc-123">Colin Hines</at> check this out"#;
        let (content, mentions) = parse_and_rewrite_mentions(input);
        assert!(content.contains(r#"itemid="0">Colin Hines</span>"#));
        assert!(content.contains("check this out"));
        let mentions_arr: Vec<serde_json::Value> =
            serde_json::from_str(mentions.as_ref().unwrap()).unwrap();
        assert_eq!(mentions_arr.len(), 1);
        assert_eq!(mentions_arr[0]["mri"], "8:orgid:abc-123");
        assert_eq!(mentions_arr[0]["displayName"], "Colin Hines");
        assert_eq!(mentions_arr[0]["mentionType"], "person");
        assert_eq!(mentions_arr[0]["itemid"], 0);
    }

    #[test]
    fn parse_mentions_multiple_mentions() {
        let input = r#"<at id="8:orgid:aaa">Alice</at> and <at id="8:orgid:bbb">Bob</at> hello"#;
        let (content, mentions) = parse_and_rewrite_mentions(input);
        assert!(content.contains(r#"itemid="0">Alice</span>"#));
        assert!(content.contains(r#"itemid="1">Bob</span>"#));
        let mentions_arr: Vec<serde_json::Value> =
            serde_json::from_str(mentions.as_ref().unwrap()).unwrap();
        assert_eq!(mentions_arr.len(), 2);
        assert_eq!(mentions_arr[0]["mri"], "8:orgid:aaa");
        assert_eq!(mentions_arr[1]["mri"], "8:orgid:bbb");
    }

    #[test]
    fn parse_mentions_ignores_non_mri_id() {
        let input = r#"<at id="0">John Doe</at> hi"#;
        let (content, mentions) = parse_and_rewrite_mentions(input);
        assert_eq!(content, r#"<at id="0">John Doe</at> hi"#);
        assert!(mentions.is_none());
    }

    #[test]
    fn parse_mentions_supports_bot_mri() {
        let input = r#"<at id="28:abcd-1234">Copilot</at> summarize"#;
        let (content, mentions) = parse_and_rewrite_mentions(input);
        assert!(content.contains(r#"itemid="0">Copilot</span>"#));
        let mentions_arr: Vec<serde_json::Value> =
            serde_json::from_str(mentions.as_ref().unwrap()).unwrap();
        assert_eq!(mentions_arr[0]["mri"], "28:abcd-1234");
    }

    #[test]
    fn parse_mentions_unescapes_display_name() {
        let input = r#"<at id="8:orgid:x">A &amp; B</at>"#;
        let (content, mentions) = parse_and_rewrite_mentions(input);
        assert!(content.contains("A & B"));
        let mentions_arr: Vec<serde_json::Value> =
            serde_json::from_str(mentions.as_ref().unwrap()).unwrap();
        assert_eq!(mentions_arr[0]["displayName"], "A & B");
    }

    #[test]
    fn parse_mentions_wraps_in_p_tag() {
        let input = r#"<at id="8:orgid:abc">Alice</at> hello"#;
        let (content, _) = parse_and_rewrite_mentions(input);
        assert!(content.starts_with("<p>"));
        assert!(content.ends_with("</p>"));
    }

    #[test]
    fn parse_mentions_preserves_existing_p_tag() {
        let input = r#"<p><at id="8:orgid:abc">Alice</at> hello</p>"#;
        let (content, _) = parse_and_rewrite_mentions(input);
        // Should not double-wrap
        assert!(!content.starts_with("<p><p>"));
    }

    #[test]
    fn html_unescape_entities() {
        assert_eq!(html_unescape("A &amp; B"), "A & B");
        assert_eq!(html_unescape("&lt;tag&gt;"), "<tag>");
        assert_eq!(html_unescape("&quot;hi&quot;"), "\"hi\"");
        assert_eq!(html_unescape("it&#39;s"), "it's");
    }
}
