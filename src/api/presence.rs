use serde::{Deserialize, Serialize};

use crate::auth::token::TokenSet;
use crate::error::{Result, TeamsError};

use super::HttpClient;

const PRESENCE_URL: &str = "https://presence.teams.microsoft.com/v1/me/forceavailability/";

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresencePayload {
    pub availability: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresenceResponse {
    pub availability: String,
}

pub struct PresenceClient<'a> {
    http: &'a HttpClient,
    tokens: &'a TokenSet,
    /// The authz-exchanged skype token (messaging token), if available.
    messaging_token: Option<&'a str>,
}

impl<'a> PresenceClient<'a> {
    pub fn new(http: &'a HttpClient, tokens: &'a TokenSet) -> Self {
        Self {
            http,
            tokens,
            messaging_token: None,
        }
    }

    pub fn with_messaging_token(
        http: &'a HttpClient,
        tokens: &'a TokenSet,
        messaging_token: &'a str,
    ) -> Self {
        Self {
            http,
            tokens,
            messaging_token: Some(messaging_token),
        }
    }

    /// Set the current user's presence/availability.
    /// Valid values: Available, Busy, DoNotDisturb, BeRightBack, Away
    pub async fn set_presence(&self, availability: &str) -> Result<()> {
        let bearer = self.tokens.skype_bearer();
        let payload = PresencePayload {
            availability: availability.to_string(),
        };

        self.http
            .execute_with_retry(|| {
                self.http
                    .client
                    .put(PRESENCE_URL)
                    .header("Authorization", &bearer)
                    .header("Content-Type", "application/json")
                    .json(&payload)
            })
            .await?;

        Ok(())
    }

    /// Clear the forced availability override, reverting presence to Offline.
    /// Tries DELETE with multiple auth strategies since the endpoint has
    /// stricter auth requirements than the PUT.
    pub async fn clear_presence(&self) -> Result<()> {
        // Strategy 1: DELETE with skypetoken= header (same auth style as messages API)
        if let Some(msg_token) = self.messaging_token {
            let skype_auth = format!("skypetoken={}", msg_token);
            let result = self
                .http
                .execute_with_retry(|| {
                    self.http
                        .client
                        .delete(PRESENCE_URL)
                        .header("Authentication", &skype_auth)
                })
                .await;
            if result.is_ok() {
                return Ok(());
            }
            tracing::debug!("DELETE with messaging skypetoken failed, trying next strategy");
        }

        // Strategy 2: DELETE with skypetoken= using the OAuth skype token
        let skype_auth = format!("skypetoken={}", self.tokens.skype.raw);
        let result = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .delete(PRESENCE_URL)
                    .header("Authentication", &skype_auth)
            })
            .await;
        if result.is_ok() {
            return Ok(());
        }
        tracing::debug!("DELETE with OAuth skypetoken failed, trying next strategy");

        // Strategy 3: DELETE with X-Skypetoken header + Bearer auth
        let bearer = self.tokens.skype_bearer();
        let result = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .delete(PRESENCE_URL)
                    .header("Authorization", &bearer)
                    .header("X-Skypetoken", &self.tokens.skype.raw)
            })
            .await;
        if result.is_ok() {
            return Ok(());
        }
        tracing::debug!("DELETE with X-Skypetoken failed, trying next strategy");

        // Strategy 4: PUT with "Away" as a workaround — not truly offline
        // but the closest achievable state without the correct DELETE auth
        tracing::warn!("all DELETE strategies failed; falling back to setting status to Away");
        self.set_presence("Away").await?;

        Ok(())
    }

    /// Get the current user's presence status.
    pub async fn get_presence(&self) -> Result<PresenceResponse> {
        let url = "https://presence.teams.microsoft.com/v1/me/presence";
        let bearer = self.tokens.skype_bearer();

        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .get(url)
                    .header("Authorization", &bearer)
            })
            .await?;

        resp.json::<PresenceResponse>()
            .await
            .map_err(|e| TeamsError::ApiError {
                status: 0,
                message: format!("failed to parse presence response: {e}"),
            })
    }
}
