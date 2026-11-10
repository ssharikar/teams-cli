use crate::auth::token::TokenSet;
use crate::error::Result;
use crate::models::{Tenant, User, UserResponse, UsersResponse, VerifiedDomain};

use super::HttpClient;

pub struct MtClient<'a> {
    http: &'a HttpClient,
    tokens: &'a TokenSet,
    region: String,
}

impl<'a> MtClient<'a> {
    pub fn new(http: &'a HttpClient, tokens: &'a TokenSet, region: &str) -> Self {
        Self {
            http,
            tokens,
            region: region.to_string(),
        }
    }

    fn base_url(&self) -> String {
        // region can be a full URL (from authz) or just a region name
        if self.region.starts_with("http") {
            format!("{}/beta", self.region)
        } else {
            format!("https://teams.microsoft.com/api/mt/{}/beta", self.region)
        }
    }

    pub async fn get_me(&self) -> Result<User> {
        // MT doesn't have a /users/me endpoint. Extract email from token and look up.
        let email = crate::auth::token::extract_username(&self.tokens.skype.raw)
            .or_else(|_| crate::auth::token::extract_username(&self.tokens.teams.raw))?;
        self.get_user(&email).await
    }

    pub async fn get_user(&self, email: &str) -> Result<User> {
        let encoded_email = urlencoding::encode(email);
        let url = format!(
            "{}/users/{}/?throwIfNotFound=false&isMailAddress=true&enableGuest=true&includeIBBarredUsers=true&skypeTeamsInfo=true",
            self.base_url(),
            encoded_email,
        );
        let bearer = self.tokens.skype_bearer();

        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        let wrapper =
            resp.json::<UserResponse>()
                .await
                .map_err(|e| crate::error::TeamsError::ApiError {
                    status: 0,
                    message: format!("failed to parse user: {e}"),
                })?;

        wrapper
            .value
            .ok_or_else(|| crate::error::TeamsError::NotFound("user not found".into()))
    }

    pub async fn fetch_short_profiles(&self, mris: &[String]) -> Result<Vec<User>> {
        let url = format!(
            "{}/users/fetchShortProfile?isMailAddress=false&enableGuest=true&includeIBBarredUsers=false&skypeTeamsInfo=true",
            self.base_url(),
        );
        let bearer = self.tokens.skype_bearer();

        let resp = self
            .http
            .execute_with_retry(|| {
                self.http
                    .client
                    .post(&url)
                    .header("Authorization", &bearer)
                    .json(&mris)
            })
            .await?;

        let users_resp =
            resp.json::<UsersResponse>()
                .await
                .map_err(|e| crate::error::TeamsError::ApiError {
                    status: 0,
                    message: format!("failed to parse users: {e}"),
                })?;

        Ok(users_resp.value)
    }

    pub async fn get_tenants(&self) -> Result<Vec<Tenant>> {
        let url = format!("{}/users/tenants", self.base_url());
        let bearer = self.tokens.skype_bearer();

        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        resp.json::<Vec<Tenant>>()
            .await
            .map_err(|e| crate::error::TeamsError::ApiError {
                status: 0,
                message: format!("failed to parse tenants: {e}"),
            })
    }

    pub async fn get_verified_domains(&self) -> Result<Vec<VerifiedDomain>> {
        let url = format!("{}/tenant/verifiedDomains", self.base_url());
        let bearer = self.tokens.skype_bearer();

        let resp = self
            .http
            .execute_with_retry(|| self.http.client.get(&url).header("Authorization", &bearer))
            .await?;

        resp.json::<Vec<VerifiedDomain>>()
            .await
            .map_err(|e| crate::error::TeamsError::ApiError {
                status: 0,
                message: format!("failed to parse domains: {e}"),
            })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::auth::token::{TokenInfo, TokenType};
    use crate::config::NetworkConfig;

    fn dummy_token_info() -> TokenInfo {
        TokenInfo {
            raw: String::new(),
            token_type: TokenType::AccessToken,
            expires_at: None,
            audience: String::new(),
        }
    }

    fn dummy_token_set() -> TokenSet {
        TokenSet {
            teams: dummy_token_info(),
            skype: dummy_token_info(),
            chatsvcagg: dummy_token_info(),
            outlook: None,
            profile: String::new(),
            tenant_id: String::new(),
        }
    }

    #[test]
    fn base_url_with_full_http_region() {
        let http = HttpClient::new(&NetworkConfig::default());
        let tokens = dummy_token_set();
        let client = MtClient::new(&http, &tokens, "https://teams.microsoft.com/api/mt/amer");
        assert_eq!(
            client.base_url(),
            "https://teams.microsoft.com/api/mt/amer/beta"
        );
    }

    #[test]
    fn base_url_with_plain_region() {
        let http = HttpClient::new(&NetworkConfig::default());
        let tokens = dummy_token_set();
        let client = MtClient::new(&http, &tokens, "amer");
        assert_eq!(
            client.base_url(),
            "https://teams.microsoft.com/api/mt/amer/beta"
        );
    }

    #[test]
    fn base_url_with_trailing_slash_region() {
        let http = HttpClient::new(&NetworkConfig::default());
        let tokens = dummy_token_set();
        let client = MtClient::new(&http, &tokens, "https://example.com/");
        // Documents current behavior: trailing slash results in "/beta" appended
        assert_eq!(client.base_url(), "https://example.com//beta");
    }
}
