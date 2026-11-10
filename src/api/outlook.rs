use crate::error::{Result, TeamsError};
use crate::models::outlook::{
    CreateEventRequest, DraftMessage, EventListResponse, MessageListResponse, OutlookEvent,
    OutlookMessage, SendMailRequest,
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
        let encoded_folder = urlencoding::encode(folder);
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
