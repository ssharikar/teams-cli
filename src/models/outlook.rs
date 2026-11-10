use serde::{Deserialize, Serialize};

// ── Email models ──────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OutlookMessage {
    #[serde(rename = "Id", default)]
    pub id: String,
    #[serde(rename = "Subject", default)]
    pub subject: String,
    #[serde(rename = "From", default)]
    pub from: Option<Recipient>,
    #[serde(rename = "ToRecipients", default)]
    pub to_recipients: Vec<Recipient>,
    #[serde(rename = "CcRecipients", default)]
    pub cc_recipients: Vec<Recipient>,
    #[serde(rename = "BccRecipients", default)]
    pub bcc_recipients: Vec<Recipient>,
    #[serde(rename = "ReceivedDateTime", default)]
    pub received_date_time: String,
    #[serde(rename = "Body", default)]
    pub body: Option<ItemBody>,
    #[serde(rename = "BodyPreview", default)]
    pub body_preview: String,
    #[serde(rename = "IsRead", default)]
    pub is_read: bool,
    #[serde(rename = "HasAttachments", default)]
    pub has_attachments: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Recipient {
    #[serde(rename = "EmailAddress")]
    pub email_address: EmailAddress,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EmailAddress {
    #[serde(rename = "Name", default)]
    pub name: String,
    #[serde(rename = "Address", default)]
    pub address: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ItemBody {
    #[serde(rename = "ContentType", default)]
    pub content_type: String,
    #[serde(rename = "Content", default)]
    pub content: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MessageListResponse {
    #[serde(rename = "value", default)]
    pub value: Vec<OutlookMessage>,
    #[serde(rename = "@odata.nextLink", default)]
    pub next_link: Option<String>,
}

// ── Mail folders ──────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MailFolder {
    #[serde(rename = "Id", default)]
    pub id: String,
    #[serde(rename = "DisplayName", default)]
    pub display_name: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MailFolderListResponse {
    #[serde(rename = "value", default)]
    pub value: Vec<MailFolder>,
}

// ── Send mail ─────────────────────────────────────────────────────

#[derive(Debug, Serialize)]
pub struct SendMailRequest {
    #[serde(rename = "Message")]
    pub message: SendMailMessage,
    #[serde(rename = "SaveToSentItems")]
    pub save_to_sent_items: bool,
}

#[derive(Debug, Serialize)]
pub struct SendMailMessage {
    #[serde(rename = "Subject")]
    pub subject: String,
    #[serde(rename = "Body")]
    pub body: ItemBody,
    #[serde(rename = "ToRecipients")]
    pub to_recipients: Vec<Recipient>,
    #[serde(rename = "CcRecipients", skip_serializing_if = "Vec::is_empty")]
    pub cc_recipients: Vec<Recipient>,
    #[serde(rename = "BccRecipients", skip_serializing_if = "Vec::is_empty")]
    pub bcc_recipients: Vec<Recipient>,
    #[serde(rename = "Attachments", skip_serializing_if = "Vec::is_empty")]
    pub attachments: Vec<FileAttachment>,
}

/// A small file attached directly in the message body as base64 (Outlook
/// REST API v2.0's inline attachment form — fine for typical attachments;
/// very large files would need the separate chunked-upload session API,
/// which this client does not implement).
#[derive(Debug, Serialize)]
pub struct FileAttachment {
    #[serde(rename = "@odata.type")]
    pub odata_type: String,
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ContentBytes")]
    pub content_bytes: String,
}

impl FileAttachment {
    pub fn new(name: String, content_bytes: String) -> Self {
        Self {
            odata_type: "#Microsoft.OutlookServices.FileAttachment".to_string(),
            name,
            content_bytes,
        }
    }
}

// ── Draft mail ────────────────────────────────────────────────────

/// Body for POST /me/messages, which creates the message in the Drafts
/// folder without sending it (sending is a separate, explicit step via
/// /messages/{id}/send — this type is never used for that).
#[derive(Debug, Serialize)]
pub struct DraftMessage {
    #[serde(rename = "Subject")]
    pub subject: String,
    #[serde(rename = "Body")]
    pub body: ItemBody,
    #[serde(rename = "ToRecipients", skip_serializing_if = "Vec::is_empty")]
    pub to_recipients: Vec<Recipient>,
    #[serde(rename = "CcRecipients", skip_serializing_if = "Vec::is_empty")]
    pub cc_recipients: Vec<Recipient>,
    #[serde(rename = "BccRecipients", skip_serializing_if = "Vec::is_empty")]
    pub bcc_recipients: Vec<Recipient>,
    #[serde(rename = "Attachments", skip_serializing_if = "Vec::is_empty")]
    pub attachments: Vec<FileAttachment>,
}

// ── Calendar models ───────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OutlookEvent {
    #[serde(rename = "Id", default)]
    pub id: String,
    #[serde(rename = "Subject", default)]
    pub subject: String,
    #[serde(rename = "Start")]
    pub start: DateTimeTimeZone,
    #[serde(rename = "End")]
    pub end: DateTimeTimeZone,
    #[serde(rename = "Location", default)]
    pub location: Option<Location>,
    #[serde(rename = "Organizer", default)]
    pub organizer: Option<Recipient>,
    #[serde(rename = "Attendees", default)]
    pub attendees: Vec<Attendee>,
    #[serde(rename = "IsAllDay", default)]
    pub is_all_day: bool,
    #[serde(rename = "IsCancelled", default)]
    pub is_cancelled: bool,
    #[serde(rename = "Body", default)]
    pub body: Option<ItemBody>,
    #[serde(rename = "IsOnlineMeeting", default)]
    pub is_online_meeting: bool,
    #[serde(rename = "OnlineMeetingUrl", default)]
    pub online_meeting_url: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DateTimeTimeZone {
    #[serde(rename = "DateTime")]
    pub date_time: String,
    #[serde(rename = "TimeZone")]
    pub time_zone: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Location {
    #[serde(rename = "DisplayName", default)]
    pub display_name: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Attendee {
    #[serde(rename = "EmailAddress")]
    pub email_address: EmailAddress,
    #[serde(rename = "Type", default)]
    pub attendee_type: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventListResponse {
    #[serde(rename = "value", default)]
    pub value: Vec<OutlookEvent>,
    #[serde(rename = "@odata.nextLink", default)]
    pub next_link: Option<String>,
}

// ── Create event ──────────────────────────────────────────────────

#[derive(Debug, Serialize)]
pub struct CreateEventRequest {
    #[serde(rename = "Subject")]
    pub subject: String,
    #[serde(rename = "Start")]
    pub start: DateTimeTimeZone,
    #[serde(rename = "End")]
    pub end: DateTimeTimeZone,
    #[serde(rename = "Location", skip_serializing_if = "Option::is_none")]
    pub location: Option<Location>,
    #[serde(rename = "Attendees", skip_serializing_if = "Vec::is_empty")]
    pub attendees: Vec<Attendee>,
    #[serde(rename = "Body", skip_serializing_if = "Option::is_none")]
    pub body: Option<ItemBody>,
    #[serde(rename = "IsOnlineMeeting", skip_serializing_if = "Option::is_none")]
    pub is_online_meeting: Option<bool>,
}

// ── Helpers ───────────────────────────────────────────────────────

impl Recipient {
    pub fn new(address: &str) -> Self {
        Self {
            email_address: EmailAddress {
                name: String::new(),
                address: address.to_string(),
            },
        }
    }

    pub fn display(&self) -> String {
        if self.email_address.name.is_empty() {
            self.email_address.address.clone()
        } else {
            format!(
                "{} <{}>",
                self.email_address.name, self.email_address.address
            )
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn deserialize_message_list_response() {
        let json = r#"{
            "value": [
                {
                    "Id": "msg-1",
                    "Subject": "Hello",
                    "ReceivedDateTime": "2026-04-24T12:00:00Z",
                    "BodyPreview": "Hi there",
                    "IsRead": true,
                    "HasAttachments": false
                }
            ]
        }"#;
        let resp: MessageListResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.value.len(), 1);
        assert_eq!(resp.value[0].id, "msg-1");
        assert_eq!(resp.value[0].subject, "Hello");
        assert!(resp.value[0].is_read);
    }

    #[test]
    fn deserialize_message_with_from() {
        let json = r#"{
            "Id": "msg-2",
            "Subject": "Test",
            "From": {
                "EmailAddress": {
                    "Name": "Alice",
                    "Address": "alice@example.com"
                }
            },
            "ReceivedDateTime": "2026-04-24T12:00:00Z",
            "BodyPreview": "preview"
        }"#;
        let msg: OutlookMessage = serde_json::from_str(json).unwrap();
        let from = msg.from.unwrap();
        assert_eq!(from.email_address.name, "Alice");
        assert_eq!(from.email_address.address, "alice@example.com");
    }

    #[test]
    fn serialize_send_mail_request() {
        let req = SendMailRequest {
            message: SendMailMessage {
                subject: "Test".to_string(),
                body: ItemBody {
                    content_type: "Text".to_string(),
                    content: "Hello".to_string(),
                },
                to_recipients: vec![Recipient::new("bob@example.com")],
                cc_recipients: vec![],
                bcc_recipients: vec![],
                attachments: vec![],
            },
            save_to_sent_items: true,
        };
        let json = serde_json::to_value(&req).unwrap();
        assert_eq!(json["Message"]["Subject"], "Test");
        assert_eq!(json["Message"]["Body"]["Content"], "Hello");
        assert!(json["Message"].get("CcRecipients").is_none());
        assert!(json["Message"].get("BccRecipients").is_none());
        assert!(json["Message"].get("Attachments").is_none());
        assert_eq!(json["SaveToSentItems"], true);
    }

    #[test]
    fn serialize_draft_message_with_bcc() {
        let draft = DraftMessage {
            subject: "Test".to_string(),
            body: ItemBody {
                content_type: "HTML".to_string(),
                content: "<p>Hi</p>".to_string(),
            },
            to_recipients: vec![Recipient::new("alice@example.com")],
            cc_recipients: vec![],
            bcc_recipients: vec![Recipient::new("bob@example.com")],
            attachments: vec![],
        };
        let json = serde_json::to_value(&draft).unwrap();
        assert_eq!(json["Subject"], "Test");
        assert!(json.get("CcRecipients").is_none());
        assert_eq!(
            json["BccRecipients"][0]["EmailAddress"]["Address"],
            "bob@example.com"
        );
    }

    #[test]
    fn serialize_draft_message_with_attachment() {
        let draft = DraftMessage {
            subject: "Test".to_string(),
            body: ItemBody {
                content_type: "Text".to_string(),
                content: "see attached".to_string(),
            },
            to_recipients: vec![Recipient::new("alice@example.com")],
            cc_recipients: vec![],
            bcc_recipients: vec![],
            attachments: vec![FileAttachment::new(
                "report.pdf".to_string(),
                "ZmFrZQ==".to_string(),
            )],
        };
        let json = serde_json::to_value(&draft).unwrap();
        assert_eq!(
            json["Attachments"][0]["@odata.type"],
            "#Microsoft.OutlookServices.FileAttachment"
        );
        assert_eq!(json["Attachments"][0]["Name"], "report.pdf");
        assert_eq!(json["Attachments"][0]["ContentBytes"], "ZmFrZQ==");
    }

    #[test]
    fn deserialize_event_list_response() {
        let json = r#"{
            "value": [
                {
                    "Id": "evt-1",
                    "Subject": "Standup",
                    "Start": {"DateTime": "2026-04-25T09:00:00", "TimeZone": "UTC"},
                    "End": {"DateTime": "2026-04-25T09:30:00", "TimeZone": "UTC"},
                    "IsAllDay": false,
                    "IsCancelled": false,
                    "IsOnlineMeeting": true,
                    "OnlineMeetingUrl": "https://teams.microsoft.com/meet/123"
                }
            ]
        }"#;
        let resp: EventListResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.value.len(), 1);
        assert_eq!(resp.value[0].subject, "Standup");
        assert!(resp.value[0].is_online_meeting);
    }

    #[test]
    fn serialize_create_event_request() {
        let req = CreateEventRequest {
            subject: "Meeting".to_string(),
            start: DateTimeTimeZone {
                date_time: "2026-04-25T10:00:00".to_string(),
                time_zone: "UTC".to_string(),
            },
            end: DateTimeTimeZone {
                date_time: "2026-04-25T11:00:00".to_string(),
                time_zone: "UTC".to_string(),
            },
            location: Some(Location {
                display_name: "Room 1".to_string(),
            }),
            attendees: vec![Attendee {
                email_address: EmailAddress {
                    name: String::new(),
                    address: "alice@example.com".to_string(),
                },
                attendee_type: "Required".to_string(),
            }],
            body: None,
            is_online_meeting: Some(true),
        };
        let json = serde_json::to_value(&req).unwrap();
        assert_eq!(json["Subject"], "Meeting");
        assert_eq!(json["Location"]["DisplayName"], "Room 1");
        assert_eq!(json["IsOnlineMeeting"], true);
        assert!(json.get("Body").is_none());
    }

    #[test]
    fn recipient_display_with_name() {
        let r = Recipient {
            email_address: EmailAddress {
                name: "Alice".to_string(),
                address: "alice@example.com".to_string(),
            },
        };
        assert_eq!(r.display(), "Alice <alice@example.com>");
    }

    #[test]
    fn recipient_display_without_name() {
        let r = Recipient::new("bob@example.com");
        assert_eq!(r.display(), "bob@example.com");
    }

    #[test]
    fn deserialize_message_missing_optional_fields() {
        let json = r#"{"Id": "msg-3"}"#;
        let msg: OutlookMessage = serde_json::from_str(json).unwrap();
        assert_eq!(msg.id, "msg-3");
        assert_eq!(msg.subject, "");
        assert!(msg.from.is_none());
        assert!(msg.to_recipients.is_empty());
        assert!(!msg.is_read);
    }
}
