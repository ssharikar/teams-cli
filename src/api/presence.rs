use serde::{Deserialize, Serialize};

use crate::auth::token::TokenSet;
use crate::error::{Result, TeamsError};

use super::HttpClient;

const PRESENCE_URL: &str = "https://presence.teams.microsoft.com/v1/me/forceavailability/";

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresencePayload {
    pub availability: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub activity: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PresenceResponse {
    pub availability: String,
}

pub struct PresenceClient<'a> {
    http: &'a HttpClient,
    tokens: &'a TokenSet,
}

impl<'a> PresenceClient<'a> {
    pub fn new(http: &'a HttpClient, tokens: &'a TokenSet) -> Self {
        Self { http, tokens }
    }

    /// Set the current user's presence/availability.
    /// For most statuses, only availability is needed. For Offline,
    /// the activity field must also be set to "OffWork".
    pub async fn set_presence(&self, availability: &str) -> Result<()> {
        let bearer = self.tokens.skype_bearer();

        let activity = match availability {
            "Offline" => Some("OffWork".to_string()),
            _ => None,
        };

        let payload = PresencePayload {
            availability: availability.to_string(),
            activity,
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
