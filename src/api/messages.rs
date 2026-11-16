use crate::error::Result;
use crate::models::{ChatMessage, MessagesResponse, SendMessageProperties, SendMessageRequest};

use super::HttpClient;

pub struct MessagesClient<'a> {
    http: &'a HttpClient,
    skype_token: &'a str,
    chat_service_url: String,
}

impl<'a> MessagesClient<'a> {
    pub fn new(http: &'a HttpClient, skype_token: &'a str, chat_service_url: &str) -> Self {
        Self {
            http,
            skype_token,
            chat_service_url: chat_service_url.to_string(),
        }
    }

    fn auth_header(&self) -> String {
        format!("skypetoken={}", self.skype_token)
    }

    pub async fn get_messages(
        &self,
        conversation_id: &str,
        page_size: u32,
    ) -> Result<Vec<ChatMessage>> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize={}&startTime=1",
            self.chat_service_url, encoded_id, page_size,
        );
        let auth = self.auth_header();

        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authentication", &auth))
            .await?;

        let messages_resp = resp.json::<MessagesResponse>().await.map_err(|e| {
            crate::error::TeamsError::ApiError {
                status: 0,
                message: format!("failed to parse messages: {e}"),
            }
        })?;

        Ok(messages_resp.messages)
    }

    pub async fn send_message(
        &self,
        conversation_id: &str,
        content: &str,
        display_name: &str,
        is_html: bool,
        mentions_json: Option<&str>,
        amsreferences: Option<Vec<String>>,
    ) -> Result<serde_json::Value> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages",
            self.chat_service_url, encoded_id,
        );
        let auth = self.auth_header();

        // Skype wire quirk: RichText/Html messages use contenttype "Text"
        // (capital T), not "text/html". Verified against real Teams client.
        let (messagetype, contenttype) = if is_html {
            ("RichText/Html", "Text")
        } else {
            ("Text", "text")
        };
        let body = SendMessageRequest {
            content: content.to_string(),
            messagetype: messagetype.to_string(),
            contenttype: contenttype.to_string(),
            clientmessageid: chrono::Utc::now().timestamp_millis().to_string(),
            imdisplayname: display_name.to_string(),
            properties: Some(SendMessageProperties {
                importance: Some(String::new()),
                subject: None,
                mentions: mentions_json.map(|s| s.to_string()),
            }),
            amsreferences,
        };

        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authentication", &auth)
                    .json(&body)
            })
            .await?;

        resp.json::<serde_json::Value>()
            .await
            .map_err(|e| crate::error::TeamsError::ApiError {
                status: 0,
                message: format!("failed to parse send response: {e}"),
            })
    }

    /// Edits the content of a message you previously sent. Uses the same
    /// request shape as `send_message` but PUT to the existing message's own
    /// URL instead of POST to the conversation's messages collection — the
    /// service records this as an edit (reflected as `properties.editTime`
    /// on later reads) rather than creating a new message.
    pub async fn edit_message(
        &self,
        conversation_id: &str,
        message_id: &str,
        content: &str,
        display_name: &str,
        is_html: bool,
        mentions_json: Option<&str>,
    ) -> Result<()> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages/{}",
            self.chat_service_url, encoded_id, message_id,
        );
        let auth = self.auth_header();

        let (messagetype, contenttype) = if is_html {
            ("RichText/Html", "Text")
        } else {
            ("Text", "text")
        };
        let body = SendMessageRequest {
            content: content.to_string(),
            messagetype: messagetype.to_string(),
            contenttype: contenttype.to_string(),
            clientmessageid: chrono::Utc::now().timestamp_millis().to_string(),
            imdisplayname: display_name.to_string(),
            properties: Some(SendMessageProperties {
                importance: Some(String::new()),
                subject: None,
                mentions: mentions_json.map(|s| s.to_string()),
            }),
            amsreferences: None,
        };

        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .put(&url)
                    .header("Authentication", &auth)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    /// Deletes a message you previously sent by setting its `deletetime`
    /// property — the same properties-PUT mechanism `react`/`unreact` use
    /// for `emotions`, just a different property name.
    pub async fn delete_message(&self, conversation_id: &str, message_id: &str) -> Result<()> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages/{}/properties?name=deletetime",
            self.chat_service_url, encoded_id, message_id,
        );
        let auth = self.auth_header();

        let timestamp = chrono::Utc::now().timestamp_millis().to_string();
        let body = serde_json::json!({
            "deletetime": timestamp,
        });

        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .put(&url)
                    .header("Authentication", &auth)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    pub async fn react(
        &self,
        conversation_id: &str,
        message_id: &str,
        reaction: &str,
    ) -> Result<()> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages/{}/properties?name=emotions",
            self.chat_service_url, encoded_id, message_id,
        );
        let auth = self.auth_header();

        let timestamp = chrono::Utc::now().timestamp_millis().to_string();
        let emotions_value = serde_json::json!({
            "key": reaction,
            "value": timestamp,
        });

        let body = serde_json::json!({
            "emotions": emotions_value.to_string(),
        });

        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .put(&url)
                    .header("Authentication", &auth)
                    .json(&body)
            })
            .await?;

        Ok(())
    }

    /// Create a new 1:1 conversation with a user by their MRI.
    /// Uses POST /v1/threads with skypetoken auth.
    /// Returns the thread ID (e.g. 19:...@unq.gbl.spaces).
    pub async fn create_conversation(
        &self,
        my_mri: &str,
        target_mri: &str,
    ) -> Result<String> {
        let url = format!("{}/v1/threads", self.chat_service_url);
        let auth = self.auth_header();

        let body = serde_json::json!({
            "members": [
                {"id": my_mri, "role": "Admin"},
                {"id": target_mri, "role": "User"}
            ],
            "properties": {
                "threadType": "chat",
                "chatFilesIndexId": "2",
                "fixedRoster": "true",
                "uniquerosterthread": "true"
            }
        });

        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authentication", &auth)
                    .json(&body)
            })
            .await?;

        // Try JSON body first for the thread ID
        let location = resp
            .headers()
            .get("location")
            .and_then(|v| v.to_str().ok())
            .map(|s| s.to_string());

        let data: serde_json::Value =
            resp.json()
                .await
                .unwrap_or_else(|_| serde_json::json!({}));

        // Check JSON id field first
        if let Some(id) = data.get("id").and_then(|v| v.as_str()) {
            return Ok(id.to_string());
        }

        // Fall back to Location header (extract thread ID from last path segment)
        if let Some(loc) = location {
            if let Some(thread_id) = loc.rsplit('/').next() {
                let decoded = urlencoding::decode(thread_id)
                    .unwrap_or_else(|_| thread_id.to_string().into());
                return Ok(decoded.to_string());
            }
        }

        Err(crate::error::TeamsError::ApiError {
            status: 0,
            message: format!(
                "conversation created but no thread ID in response: {}",
                serde_json::to_string(&data).unwrap_or_default()
            ),
        })
    }

    pub async fn unreact(
        &self,
        conversation_id: &str,
        message_id: &str,
        reaction: &str,
    ) -> Result<()> {
        let encoded_id = urlencoding::encode(conversation_id);
        let url = format!(
            "{}/v1/users/ME/conversations/{}/messages/{}/properties?name=emotions",
            self.chat_service_url, encoded_id, message_id,
        );
        let auth = self.auth_header();

        let emotions_value = serde_json::json!({
            "key": reaction,
            "value": "",
        });

        let body = serde_json::json!({
            "emotions": emotions_value.to_string(),
        });

        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .put(&url)
                    .header("Authentication", &auth)
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
    fn auth_header_format() {
        let http = HttpClient::new(&NetworkConfig::default());
        let client = MessagesClient::new(&http, "test-token-123", "https://chat.example.com");
        assert_eq!(client.auth_header(), "skypetoken=test-token-123");
    }

    #[test]
    fn conversation_id_is_url_encoded_in_get_messages_url() {
        // Verify the URL encoding logic by checking the encoded form of a typical conversation ID
        let conv_id = "19:abc@thread.v2";
        let encoded = urlencoding::encode(conv_id);
        assert_eq!(encoded, "19%3Aabc%40thread.v2");
    }

    #[test]
    fn conversation_id_encoding_preserves_safe_chars() {
        let conv_id = "simple-id-123";
        let encoded = urlencoding::encode(conv_id);
        assert_eq!(encoded, "simple-id-123");
    }

    #[test]
    fn messages_client_stores_chat_service_url() {
        let http = HttpClient::new(&NetworkConfig::default());
        let client = MessagesClient::new(&http, "tok", "https://amer.ng.msg.teams.microsoft.com");
        assert_eq!(
            client.chat_service_url,
            "https://amer.ng.msg.teams.microsoft.com"
        );
    }
}
