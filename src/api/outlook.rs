use crate::error::{Result, TeamsError};
use crate::models::outlook::{
    AttachmentListResponse, CreateEventRequest, DraftMessage, EventListResponse, MailAttachment,
    MailFolderListResponse, MessageListResponse, OutlookEvent, OutlookMessage, Recipient,
    SendMailRequest, UpdateEventRequest,
};

use super::HttpClient;

const OUTLOOK_API_BASE: &str = "https://outlook.office365.com/api/v2.0/me";

pub struct OutlookClient<'a> {
    http: &'a HttpClient,
    bearer: String,
}

impl<'a> OutlookClient<'a> {
    pub fn new(http: &'a HttpClient, bearer: String) -> Self {
        Self { http, bearer }
    }

    // ── Email ─────────────────────────────────────────────────────

    pub async fn list_messages(
        &self,
        folder: &str,
        filter: Option<&str>,
        top: u32,
    ) -> Result<Vec<OutlookMessage>> {
        let folder_id = self.resolve_folder_id(folder).await?;
        let encoded_folder = urlencoding::encode(&folder_id);
        let mut url = format!(
            "{}/mailfolders/{}/messages?$top={}&$orderby=ReceivedDateTime%20desc&$select=Id,Subject,From,ToRecipients,CcRecipients,BccRecipients,ReceivedDateTime,BodyPreview,IsRead,HasAttachments",
            OUTLOOK_API_BASE, encoded_folder, top,
        );

        if let Some(f) = filter {
            url.push_str(&format!("&$filter={}", urlencoding::encode(f)));
        }

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        let list: MessageListResponse = resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse message list: {e}"),
        })?;

        Ok(list.value)
    }

    /// Resolves a mail folder to the ID that `/mailfolders/{id}/messages`
    /// needs. Graph/Outlook accept a handful of well-known folder names
    /// (Inbox, Drafts, SentItems, ...) directly in that path, but any other
    /// folder — e.g. a custom folder like "github-access" — needs its real
    /// folder ID first, or the API fails with `ErrorInvalidIdMalformed`.
    /// Looks among top-level folders first, then Inbox's child folders,
    /// since custom folders are most commonly created there.
    async fn resolve_folder_id(&self, folder: &str) -> Result<String> {
        const WELL_KNOWN: &[&str] = &[
            "inbox",
            "drafts",
            "sentitems",
            "deleteditems",
            "outbox",
            "junkemail",
            "archive",
            "clutter",
            "conversationhistory",
        ];
        let normalized: String = folder
            .chars()
            .filter(|c| !c.is_whitespace())
            .collect::<String>()
            .to_ascii_lowercase();
        if WELL_KNOWN.contains(&normalized.as_str()) {
            return Ok(folder.to_string());
        }

        if let Some(id) = self
            .find_folder_by_name(&format!("{OUTLOOK_API_BASE}/mailfolders"), folder)
            .await?
        {
            return Ok(id);
        }
        if let Some(id) = self
            .find_folder_by_name(
                &format!("{OUTLOOK_API_BASE}/mailfolders/inbox/childfolders"),
                folder,
            )
            .await?
        {
            return Ok(id);
        }

        Err(TeamsError::NotFound(format!(
            "mail folder {folder:?} (checked top-level folders and Inbox subfolders)"
        )))
    }

    async fn find_folder_by_name(&self, folders_url: &str, name: &str) -> Result<Option<String>> {
        let odata_value = name.replace('\'', "''");
        let url = format!(
            "{folders_url}?$filter={}&$select=Id,DisplayName",
            urlencoding::encode(&format!("DisplayName eq '{odata_value}'")),
        );

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        let list: MailFolderListResponse = resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse mail folder list: {e}"),
        })?;

        Ok(list.value.into_iter().next().map(|f| f.id))
    }

    pub async fn get_message(&self, message_id: &str) -> Result<OutlookMessage> {
        let encoded_id = urlencoding::encode(message_id);
        let url = format!(
            "{}/messages/{}?$select=Id,Subject,From,ToRecipients,CcRecipients,BccRecipients,ReceivedDateTime,Body,BodyPreview,IsRead,HasAttachments",
            OUTLOOK_API_BASE, encoded_id,
        );

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse message: {e}"),
        })
    }

    pub async fn send_message(&self, request: &SendMailRequest) -> Result<()> {
        let url = format!("{}/sendmail", OUTLOOK_API_BASE);

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(request)
            })
            .await?;

        let status = resp.status();
        if status.is_success() {
            Ok(())
        } else {
            let body = resp.text().await.unwrap_or_default();
            Err(TeamsError::ApiError {
                status: status.as_u16(),
                message: format!("send mail failed: {body}"),
            })
        }
    }

    /// Creates a message in the Drafts folder via POST /me/messages. Unlike
    /// `send_message` (POST /me/sendmail), this never delivers anything —
    /// Outlook only sends a message that was created this way if a separate,
    /// explicit /messages/{id}/send call is made, which this client does not
    /// implement.
    pub async fn create_draft(&self, draft: &DraftMessage) -> Result<OutlookMessage> {
        let url = format!("{}/messages", OUTLOOK_API_BASE);

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(draft)
            })
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse created draft: {e}"),
        })
    }

    /// Sends an existing draft (e.g. one created by `create_draft`) via
    /// POST /me/messages/{id}/send. This is the only way this client ever
    /// delivers a message that was saved as a draft — `create_draft` itself
    /// never sends anything.
    pub async fn send_draft(&self, message_id: &str) -> Result<()> {
        let encoded_id = urlencoding::encode(message_id);
        let url = format!("{}/messages/{}/send", OUTLOOK_API_BASE, encoded_id);

        let bearer = self.bearer.clone();
        self.http
            .execute_with_retry(|| self.http.client.post(&url).header("Authorization", &bearer))
            .await?;

        Ok(())
    }

    /// Permanently deletes a message (e.g. a draft) by id.
    pub async fn delete_message(&self, message_id: &str) -> Result<()> {
        let encoded_id = urlencoding::encode(message_id);
        let url = format!("{}/messages/{}", OUTLOOK_API_BASE, encoded_id);

        let bearer = self.bearer.clone();
        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .delete(&url)
                    .header("Authorization", &bearer)
            })
            .await?;

        Ok(())
    }

    pub async fn search_messages(&self, query: &str, top: u32) -> Result<Vec<OutlookMessage>> {
        let encoded_query = urlencoding::encode(query);
        let url = format!(
            "{}/messages?$search=%22{}%22&$top={}&$select=Id,Subject,From,ReceivedDateTime,BodyPreview,IsRead,HasAttachments",
            OUTLOOK_API_BASE, encoded_query, top,
        );

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        let list: MessageListResponse = resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse search results: {e}"),
        })?;

        Ok(list.value)
    }

    /// Replies to a message via POST /me/messages/{id}/reply (or replyall),
    /// which sends immediately — Outlook builds the quoted-original body and
    /// keeps the thread/subject, prepending `comment` as the new text.
    pub async fn reply_to_message(
        &self,
        message_id: &str,
        comment: &str,
        reply_all: bool,
    ) -> Result<()> {
        let encoded_id = urlencoding::encode(message_id);
        let action = if reply_all { "replyall" } else { "reply" };
        let url = format!("{OUTLOOK_API_BASE}/messages/{encoded_id}/{action}");

        let bearer = self.bearer.clone();
        let body = serde_json::json!({ "Comment": comment });
        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    /// Forwards a message via POST /me/messages/{id}/forward, which sends
    /// immediately to the given recipients with `comment` prepended.
    pub async fn forward_message(
        &self,
        message_id: &str,
        comment: &str,
        to: Vec<Recipient>,
    ) -> Result<()> {
        let encoded_id = urlencoding::encode(message_id);
        let url = format!("{OUTLOOK_API_BASE}/messages/{encoded_id}/forward");

        let bearer = self.bearer.clone();
        let body = serde_json::json!({ "Comment": comment, "ToRecipients": to });
        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    /// Lists a message's attachments via GET /me/messages/{id}/attachments.
    /// The v2.0 API includes `ContentBytes` for file attachments directly in
    /// this list response, so a download rarely needs the per-attachment
    /// fetch below — but callers should still check `content_bytes` and
    /// fall back to `get_attachment` if it's missing.
    pub async fn list_attachments(&self, message_id: &str) -> Result<Vec<MailAttachment>> {
        let encoded_id = urlencoding::encode(message_id);
        let url = format!("{OUTLOOK_API_BASE}/messages/{encoded_id}/attachments");

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        let list: AttachmentListResponse = resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse attachment list: {e}"),
        })?;

        Ok(list.value)
    }

    /// Fetches a single attachment by id via GET
    /// /me/messages/{id}/attachments/{attachment_id} — used when the list
    /// response didn't include `ContentBytes`.
    pub async fn get_attachment(
        &self,
        message_id: &str,
        attachment_id: &str,
    ) -> Result<MailAttachment> {
        let encoded_msg_id = urlencoding::encode(message_id);
        let encoded_att_id = urlencoding::encode(attachment_id);
        let url =
            format!("{OUTLOOK_API_BASE}/messages/{encoded_msg_id}/attachments/{encoded_att_id}");

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse attachment: {e}"),
        })
    }

    // ── Calendar ──────────────────────────────────────────────────

    pub async fn list_events(&self, start: &str, end: &str, top: u32) -> Result<Vec<OutlookEvent>> {
        let url = format!(
            "{}/calendarview?startDateTime={}&endDateTime={}&$top={}&$orderby=Start/DateTime&$select=Id,Subject,Start,End,Location,Organizer,Attendees,IsAllDay,IsCancelled,IsOnlineMeeting,OnlineMeetingUrl",
            OUTLOOK_API_BASE,
            urlencoding::encode(start),
            urlencoding::encode(end),
            top,
        );

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .get(&url)
                    .header("Authorization", &bearer)
                    .header("Prefer", "outlook.timezone=\"UTC\"")
            })
            .await?;

        let list: EventListResponse = resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse event list: {e}"),
        })?;

        Ok(list.value)
    }

    pub async fn get_event(&self, event_id: &str) -> Result<OutlookEvent> {
        let encoded_id = urlencoding::encode(event_id);
        let url = format!("{}/events/{}", OUTLOOK_API_BASE, encoded_id);

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .get(&url)
                    .header("Authorization", &bearer)
                    .header("Prefer", "outlook.timezone=\"UTC\"")
            })
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse event: {e}"),
        })
    }

    pub async fn create_event(&self, request: &CreateEventRequest) -> Result<OutlookEvent> {
        let url = format!("{}/events", OUTLOOK_API_BASE);

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .header("Prefer", "outlook.timezone=\"UTC\"")
                    .json(request)
            })
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse created event: {e}"),
        })
    }

    /// Partially updates an event via PATCH /me/events/{id} — only fields
    /// set on `request` are changed.
    pub async fn update_event(
        &self,
        event_id: &str,
        request: &UpdateEventRequest,
    ) -> Result<OutlookEvent> {
        let encoded_id = urlencoding::encode(event_id);
        let url = format!("{OUTLOOK_API_BASE}/events/{encoded_id}");

        let bearer = self.bearer.clone();
        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .patch(&url)
                    .header("Authorization", &bearer)
                    .header("Prefer", "outlook.timezone=\"UTC\"")
                    .json(request)
            })
            .await?;

        resp.json().await.map_err(|e| TeamsError::ApiError {
            status: 0,
            message: format!("failed to parse updated event: {e}"),
        })
    }

    /// Cancels an event you organize via POST /me/events/{id}/cancel, which
    /// notifies attendees and removes it from their calendars.
    pub async fn cancel_event(&self, event_id: &str, comment: Option<&str>) -> Result<()> {
        let encoded_id = urlencoding::encode(event_id);
        let url = format!("{OUTLOOK_API_BASE}/events/{encoded_id}/cancel");

        let bearer = self.bearer.clone();
        let body = serde_json::json!({ "Comment": comment.unwrap_or("") });
        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    /// Responds to an invite you're attending, via POST
    /// /me/events/{id}/accept, /decline, or /tentativelyaccept.
    pub async fn respond_to_event(
        &self,
        event_id: &str,
        response: &str,
        comment: Option<&str>,
    ) -> Result<()> {
        let action = match response {
            "accept" => "accept",
            "decline" => "decline",
            "tentative" => "tentativelyaccept",
            other => {
                return Err(TeamsError::InvalidInput(format!(
                    "unknown response {other:?} (use accept, decline, or tentative)"
                )))
            }
        };
        let encoded_id = urlencoding::encode(event_id);
        let url = format!("{OUTLOOK_API_BASE}/events/{encoded_id}/{action}");

        let bearer = self.bearer.clone();
        let body = serde_json::json!({ "Comment": comment.unwrap_or("") });
        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(&body)
            })
            .await?;

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::NetworkConfig;

    #[test]
    fn outlook_client_stores_bearer() {
        let http = HttpClient::new(&NetworkConfig::default());
        let client = OutlookClient::new(&http, "Bearer test-token".to_string());
        assert_eq!(client.bearer, "Bearer test-token");
    }

    #[test]
    fn outlook_api_base_is_https() {
        assert!(OUTLOOK_API_BASE.starts_with("https://"));
    }
}
