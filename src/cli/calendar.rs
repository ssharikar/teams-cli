use clap::{Args, Subcommand};
use std::time::Instant;

use crate::api::outlook::OutlookClient;
use crate::api::HttpClient;
use crate::auth::token::TokenSet;
use crate::error::{Result, TeamsError};
use crate::html::strip_html;
use crate::models::outlook::{
    Attendee, CreateEventRequest, DateTimeTimeZone, EmailAddress, ItemBody, Location,
    UpdateEventRequest,
};
use crate::output::{self, OutputFormat};

#[derive(Args)]
pub struct CalendarArgs {
    #[command(subcommand)]
    pub command: CalendarCommand,
}

#[derive(Subcommand)]
pub enum CalendarCommand {
    /// List upcoming events
    List {
        /// Start date/time (ISO 8601 or relative: "today", "tomorrow", "now")
        #[arg(long)]
        from: Option<String>,
        /// End date/time (ISO 8601 or relative: "today", "tomorrow", "+3d", "+1w")
        #[arg(long)]
        to: Option<String>,
        /// Maximum number of events
        #[arg(long, default_value = "25")]
        limit: u32,
    },
    /// Get event details
    Get {
        /// Event ID
        event_id: String,
    },
    /// Create a calendar event or meeting
    Create {
        /// Event subject/title
        #[arg(long, required = true)]
        subject: String,
        /// Start date/time (ISO 8601, e.g. "2026-04-25T10:00:00")
        #[arg(long, required = true)]
        start: String,
        /// End date/time (ISO 8601)
        #[arg(long, required = true)]
        end: String,
        /// Time zone (default: UTC)
        #[arg(long, default_value = "UTC")]
        timezone: String,
        /// Location
        #[arg(long)]
        location: Option<String>,
        /// Attendee email addresses (repeatable)
        #[arg(long)]
        attendees: Vec<String>,
        /// Create as online meeting (Teams meeting)
        #[arg(long)]
        online: bool,
        /// Event body/description
        #[arg(long)]
        body: Option<String>,
    },
    /// Update fields on an existing event (only given flags are changed)
    Update {
        /// Event ID
        event_id: String,
        /// New subject/title
        #[arg(long)]
        subject: Option<String>,
        /// New start date/time (ISO 8601)
        #[arg(long)]
        start: Option<String>,
        /// New end date/time (ISO 8601)
        #[arg(long)]
        end: Option<String>,
        /// Time zone for --start/--end (default: UTC)
        #[arg(long, default_value = "UTC")]
        timezone: String,
        /// New location
        #[arg(long)]
        location: Option<String>,
        /// New body/description
        #[arg(long)]
        body: Option<String>,
    },
    /// Cancel an event you organize, notifying attendees
    Cancel {
        /// Event ID
        event_id: String,
        /// Optional cancellation message to attendees
        #[arg(long)]
        comment: Option<String>,
    },
    /// Respond to an event you're invited to
    Respond {
        /// Event ID
        event_id: String,
        /// Response: accept, decline, or tentative
        #[arg(long, required = true)]
        response: String,
        /// Optional message to the organizer
        #[arg(long)]
        comment: Option<String>,
    },
}

pub async fn handle(
    args: &CalendarArgs,
    tokens: &TokenSet,
    http: &HttpClient,
    format: OutputFormat,
) -> Result<()> {
    let bearer = tokens.outlook_bearer()?;
    let client = OutlookClient::new(http, bearer);

    match &args.command {
        CalendarCommand::List { from, to, limit } => {
            let start_dt = match from {
                Some(s) => parse_datetime(s)?,
                None => chrono::Utc::now().to_rfc3339(),
            };
            let end_dt = match to {
                Some(s) => parse_datetime(s)?,
                None => {
                    let end = chrono::Utc::now() + chrono::Duration::days(7);
                    end.to_rfc3339()
                }
            };

            let start = Instant::now();
            let events = client.list_events(&start_dt, &end_dt, *limit).await?;

            let display: Vec<serde_json::Value> = events
                .iter()
                .map(|e| {
                    let location = e
                        .location
                        .as_ref()
                        .map(|l| l.display_name.as_str())
                        .unwrap_or("");
                    let organizer = e
                        .organizer
                        .as_ref()
                        .map(|r| r.display())
                        .unwrap_or_default();
                    serde_json::json!({
                        "id": e.id,
                        "subject": e.subject,
                        "start": e.start.date_time,
                        "end": e.end.date_time,
                        "timezone": e.start.time_zone,
                        "location": location,
                        "organizer": organizer,
                        "all_day": e.is_all_day,
                        "cancelled": e.is_cancelled,
                        "online": e.is_online_meeting,
                        "meeting_url": e.online_meeting_url,
                    })
                })
                .collect();

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        CalendarCommand::Get { event_id } => {
            let start = Instant::now();
            let event = client.get_event(event_id).await?;

            let attendees: Vec<serde_json::Value> = event
                .attendees
                .iter()
                .map(|a| {
                    serde_json::json!({
                        "email": a.email_address.address,
                        "name": a.email_address.name,
                        "type": a.attendee_type,
                    })
                })
                .collect();

            let body_text = event
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

            let display = serde_json::json!({
                "id": event.id,
                "subject": event.subject,
                "start": event.start.date_time,
                "end": event.end.date_time,
                "timezone": event.start.time_zone,
                "location": event.location.as_ref().map(|l| &l.display_name),
                "organizer": event.organizer.as_ref().map(|r| r.display()),
                "attendees": attendees,
                "all_day": event.is_all_day,
                "cancelled": event.is_cancelled,
                "online": event.is_online_meeting,
                "meeting_url": event.online_meeting_url,
                "body": body_text,
            });

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        CalendarCommand::Create {
            subject,
            start: start_str,
            end: end_str,
            timezone,
            location,
            attendees,
            online,
            body,
        } => {
            let request = CreateEventRequest {
                subject: subject.clone(),
                start: DateTimeTimeZone {
                    date_time: start_str.clone(),
                    time_zone: timezone.clone(),
                },
                end: DateTimeTimeZone {
                    date_time: end_str.clone(),
                    time_zone: timezone.clone(),
                },
                location: location.as_ref().map(|l| Location {
                    display_name: l.clone(),
                }),
                attendees: attendees
                    .iter()
                    .map(|addr| Attendee {
                        email_address: EmailAddress {
                            name: String::new(),
                            address: addr.clone(),
                        },
                        attendee_type: "Required".to_string(),
                    })
                    .collect(),
                body: body.as_ref().map(|b| ItemBody {
                    content_type: "Text".to_string(),
                    content: b.clone(),
                }),
                is_online_meeting: if *online { Some(true) } else { None },
            };

            let start = Instant::now();
            let event = client.create_event(&request).await?;

            let display = serde_json::json!({
                "id": event.id,
                "subject": event.subject,
                "start": event.start.date_time,
                "end": event.end.date_time,
                "online": event.is_online_meeting,
                "meeting_url": event.online_meeting_url,
            });

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        CalendarCommand::Update {
            event_id,
            subject,
            start: start_str,
            end: end_str,
            timezone,
            location,
            body,
        } => {
            if subject.is_none()
                && start_str.is_none()
                && end_str.is_none()
                && location.is_none()
                && body.is_none()
            {
                return Err(TeamsError::InvalidInput(
                    "provide at least one of --subject, --start, --end, --location, --body".into(),
                ));
            }

            let request = UpdateEventRequest {
                subject: subject.clone(),
                start: match start_str {
                    Some(s) => Some(DateTimeTimeZone {
                        date_time: parse_datetime(s)?,
                        time_zone: timezone.clone(),
                    }),
                    None => None,
                },
                end: match end_str {
                    Some(s) => Some(DateTimeTimeZone {
                        date_time: parse_datetime(s)?,
                        time_zone: timezone.clone(),
                    }),
                    None => None,
                },
                location: location.as_ref().map(|l| Location {
                    display_name: l.clone(),
                }),
                body: body.as_ref().map(|b| ItemBody {
                    content_type: "Text".to_string(),
                    content: b.clone(),
                }),
            };

            let start = Instant::now();
            let event = client.update_event(event_id, &request).await?;

            let display = serde_json::json!({
                "id": event.id,
                "subject": event.subject,
                "start": event.start.date_time,
                "end": event.end.date_time,
                "location": event.location.as_ref().map(|l| &l.display_name),
            });

            output::print_output(format, display, start.elapsed().as_millis() as u64);
        }

        CalendarCommand::Cancel { event_id, comment } => {
            let start = Instant::now();
            client.cancel_event(event_id, comment.as_deref()).await?;

            let result = serde_json::json!({
                "status": "cancelled",
                "id": event_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }

        CalendarCommand::Respond {
            event_id,
            response,
            comment,
        } => {
            let start = Instant::now();
            client
                .respond_to_event(event_id, response, comment.as_deref())
                .await?;

            let result = serde_json::json!({
                "status": "responded",
                "response": response,
                "id": event_id,
            });
            output::print_output(format, result, start.elapsed().as_millis() as u64);
        }
    }

    Ok(())
}

fn parse_datetime(input: &str) -> Result<String> {
    let s = input.trim().to_lowercase();

    let now = chrono::Utc::now();

    match s.as_str() {
        "now" => return Ok(now.to_rfc3339()),
        "today" => {
            let today = now.date_naive().and_hms_opt(0, 0, 0).unwrap();
            return Ok(chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(
                today,
                chrono::Utc,
            )
            .to_rfc3339());
        }
        "tomorrow" => {
            let tomorrow = (now.date_naive() + chrono::Duration::days(1))
                .and_hms_opt(0, 0, 0)
                .unwrap();
            return Ok(chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(
                tomorrow,
                chrono::Utc,
            )
            .to_rfc3339());
        }
        _ => {}
    }

    // Relative offsets: +3d, +1w, +2h
    if s.starts_with('+') && s.len() >= 3 {
        let (num_str, unit) = s[1..].split_at(s.len() - 2);
        if let Ok(num) = num_str.parse::<i64>() {
            let duration = match unit {
                "h" => chrono::Duration::hours(num),
                "d" => chrono::Duration::days(num),
                "w" => chrono::Duration::weeks(num),
                _ => {
                    return Err(TeamsError::InvalidInput(format!(
                        "invalid relative offset '{input}': expected +Nh, +Nd, or +Nw"
                    )))
                }
            };
            return Ok((now + duration).to_rfc3339());
        }
    }

    // Try ISO 8601 parse
    if chrono::DateTime::parse_from_rfc3339(input).is_ok() {
        return Ok(input.to_string());
    }

    // Try naive datetime (no timezone, assume UTC)
    if let Ok(naive) = chrono::NaiveDateTime::parse_from_str(input, "%Y-%m-%dT%H:%M:%S") {
        return Ok(
            chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(naive, chrono::Utc)
                .to_rfc3339(),
        );
    }

    // Try date-only
    if let Ok(date) = chrono::NaiveDate::parse_from_str(input, "%Y-%m-%d") {
        let dt = date.and_hms_opt(0, 0, 0).unwrap();
        return Ok(
            chrono::DateTime::<chrono::Utc>::from_naive_utc_and_offset(dt, chrono::Utc)
                .to_rfc3339(),
        );
    }

    Err(TeamsError::InvalidInput(format!(
        "could not parse date/time '{input}': expected ISO 8601, relative (+3d, +1w), or keyword (now, today, tomorrow)"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_datetime_now() {
        let result = parse_datetime("now").unwrap();
        assert!(chrono::DateTime::parse_from_rfc3339(&result).is_ok());
    }

    #[test]
    fn parse_datetime_today() {
        let result = parse_datetime("today").unwrap();
        let dt = chrono::DateTime::parse_from_rfc3339(&result).unwrap();
        assert_eq!(dt.time(), chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap());
    }

    #[test]
    fn parse_datetime_tomorrow() {
        let result = parse_datetime("tomorrow").unwrap();
        let dt = chrono::DateTime::parse_from_rfc3339(&result).unwrap();
        let expected_date = (chrono::Utc::now() + chrono::Duration::days(1)).date_naive();
        assert_eq!(dt.date_naive(), expected_date);
    }

    #[test]
    fn parse_datetime_relative_days() {
        let result = parse_datetime("+3d").unwrap();
        let dt = chrono::DateTime::parse_from_rfc3339(&result).unwrap();
        let diff = dt.signed_duration_since(chrono::Utc::now());
        assert!(diff.num_days() >= 2 && diff.num_days() <= 3);
    }

    #[test]
    fn parse_datetime_relative_weeks() {
        let result = parse_datetime("+1w").unwrap();
        let dt = chrono::DateTime::parse_from_rfc3339(&result).unwrap();
        let diff = dt.signed_duration_since(chrono::Utc::now());
        assert!(diff.num_days() >= 6 && diff.num_days() <= 7);
    }

    #[test]
    fn parse_datetime_iso8601() {
        let input = "2026-04-25T10:00:00+00:00";
        let result = parse_datetime(input).unwrap();
        assert_eq!(result, input);
    }

    #[test]
    fn parse_datetime_naive() {
        let result = parse_datetime("2026-04-25T10:00:00").unwrap();
        assert!(chrono::DateTime::parse_from_rfc3339(&result).is_ok());
    }

    #[test]
    fn parse_datetime_date_only() {
        let result = parse_datetime("2026-04-25").unwrap();
        let dt = chrono::DateTime::parse_from_rfc3339(&result).unwrap();
        assert_eq!(dt.time(), chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap());
    }

    #[test]
    fn parse_datetime_invalid() {
        assert!(parse_datetime("not-a-date").is_err());
    }
}
