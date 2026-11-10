use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

use crate::error::{Result, TeamsError};

pub const TEAMS_APP_ID: &str = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
pub const SKYPE_RESOURCE: &str = "https://api.spaces.skype.com";
pub const CHATSVCAGG_RESOURCE: &str = "https://chatsvcagg.teams.microsoft.com";
pub const OUTLOOK_RESOURCE: &str = "https://outlook.office365.com";
#[allow(dead_code)]
pub const REDIRECT_URI: &str = "https://teams.microsoft.com/go";
#[allow(dead_code)]
pub const MICROSOFT_TENANT_ID: &str = "f8cdef31-a31e-4b4a-93e4-5f571e91255a";

#[derive(Clone, Serialize, Deserialize)]
pub struct TokenSet {
    pub teams: TokenInfo,
    pub skype: TokenInfo,
    pub chatsvcagg: TokenInfo,
    #[serde(default)]
    pub outlook: Option<TokenInfo>,
    pub profile: String,
    pub tenant_id: String,
}

impl std::fmt::Debug for TokenSet {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TokenSet")
            .field("teams", &self.teams)
            .field("skype", &self.skype)
            .field("chatsvcagg", &self.chatsvcagg)
            .field("outlook", &self.outlook)
            .field("profile", &self.profile)
            .field("tenant_id", &self.tenant_id)
            .finish()
    }
}

#[derive(Clone, Serialize, Deserialize)]
pub struct TokenInfo {
    pub raw: String,
    pub token_type: TokenType,
    pub expires_at: Option<DateTime<Utc>>,
    pub audience: String,
}

impl std::fmt::Debug for TokenInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TokenInfo")
            .field("raw", &"[REDACTED]")
            .field("token_type", &self.token_type)
            .field("expires_at", &self.expires_at)
            .field("audience", &self.audience)
            .finish()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TokenType {
    IdToken,
    AccessToken,
}

impl TokenInfo {
    pub fn from_jwt(raw: &str, token_type: TokenType) -> Result<Self> {
        let claims = decode_jwt_claims(raw)?;
        let expires_at = claims
            .exp
            .map(|exp| DateTime::<Utc>::from_timestamp(exp, 0).unwrap_or_else(Utc::now));
        let audience = claims.aud.unwrap_or_default();

        Ok(Self {
            raw: raw.to_string(),
            token_type,
            expires_at,
            audience,
        })
    }

    pub fn is_expired(&self) -> bool {
        match self.expires_at {
            Some(exp) => Utc::now() >= exp,
            None => true,
        }
    }
}

#[allow(dead_code)]
impl TokenSet {
    pub fn is_expired(&self) -> bool {
        self.teams.is_expired() || self.skype.is_expired() || self.chatsvcagg.is_expired()
    }

    pub fn skype_header(&self) -> String {
        format!("skypetoken={}", self.skype.raw)
    }

    pub fn chatsvcagg_bearer(&self) -> String {
        format!("Bearer {}", self.chatsvcagg.raw)
    }

    pub fn teams_bearer(&self) -> String {
        format!("Bearer {}", self.teams.raw)
    }

    pub fn skype_bearer(&self) -> String {
        format!("Bearer {}", self.skype.raw)
    }

    pub fn outlook_bearer(&self) -> crate::error::Result<String> {
        match &self.outlook {
            Some(t) => Ok(format!("Bearer {}", t.raw)),
            None => Err(TeamsError::AuthError(
                "outlook token not available; run a mail or calendar command to trigger authentication".into(),
            )),
        }
    }
}

#[derive(Debug, Deserialize)]
#[allow(dead_code)]
pub(crate) struct JwtClaims {
    #[serde(default)]
    exp: Option<i64>,
    #[serde(default)]
    aud: Option<String>,
    #[serde(default)]
    pub tid: Option<String>,
    #[serde(default)]
    pub upn: Option<String>,
    #[serde(default)]
    pub preferred_username: Option<String>,
    #[serde(default)]
    pub unique_name: Option<String>,
}

// Note: JWT signature validation is intentionally skipped. Tokens are obtained
// directly from Microsoft's OAuth endpoints and stored locally. Full JWKS
// validation would require fetching Microsoft's signing keys on every invocation.
pub(crate) fn decode_jwt_claims(token: &str) -> Result<JwtClaims> {
    let parts: Vec<&str> = token.split('.').collect();
    if parts.len() < 2 {
        return Err(TeamsError::AuthError("invalid JWT format".into()));
    }

    let payload = base64::engine::general_purpose::URL_SAFE_NO_PAD
        .decode(parts[1])
        .or_else(|_| base64::engine::general_purpose::STANDARD.decode(parts[1]))
        .map_err(|e| TeamsError::AuthError(format!("failed to decode JWT payload: {e}")))?;

    serde_json::from_slice(&payload)
        .map_err(|e| TeamsError::AuthError(format!("failed to parse JWT claims: {e}")))
}

pub fn extract_tenant_id(token: &str) -> Result<Option<String>> {
    let claims = decode_jwt_claims(token)?;
    Ok(claims.tid)
}

pub fn extract_username(token: &str) -> Result<String> {
    let claims = decode_jwt_claims(token)?;
    claims
        .upn
        .or(claims.preferred_username)
        .or(claims.unique_name)
        .ok_or_else(|| TeamsError::AuthError("no username claim in token".into()))
}

use base64::Engine;

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn make_test_jwt(claims: &serde_json::Value) -> String {
        use base64::Engine;
        let header = base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(r#"{"alg":"none"}"#);
        let payload = base64::engine::general_purpose::URL_SAFE_NO_PAD
            .encode(serde_json::to_string(claims).unwrap());
        format!("{header}.{payload}.sig")
    }

    // ── decode_jwt_claims ──────────────────────────────────────────

    #[test]
    fn decode_jwt_claims_valid() {
        let jwt = make_test_jwt(&json!({"exp": 1700000000, "aud": "https://example.com"}));
        let claims = decode_jwt_claims(&jwt).unwrap();
        assert_eq!(claims.exp, Some(1700000000));
        assert_eq!(claims.aud.as_deref(), Some("https://example.com"));
    }

    #[test]
    fn decode_jwt_claims_single_part_errors() {
        let result = decode_jwt_claims("not-a-jwt");
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("invalid JWT format"), "got: {err}");
    }

    #[test]
    fn decode_jwt_claims_invalid_base64_payload() {
        let jwt = "header.!!!invalid-base64!!!.sig";
        let result = decode_jwt_claims(jwt);
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("failed to decode JWT payload"), "got: {err}");
    }

    #[test]
    fn decode_jwt_claims_valid_base64_but_invalid_json() {
        use base64::Engine;
        let payload =
            base64::engine::general_purpose::URL_SAFE_NO_PAD.encode("this is not json at all");
        let jwt = format!("header.{payload}.sig");
        let result = decode_jwt_claims(&jwt);
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("failed to parse JWT claims"), "got: {err}");
    }

    #[test]
    fn decode_jwt_claims_url_safe_no_pad_encoding() {
        // URL_SAFE_NO_PAD is the primary encoding path
        let jwt = make_test_jwt(&json!({"exp": 42}));
        let claims = decode_jwt_claims(&jwt).unwrap();
        assert_eq!(claims.exp, Some(42));
    }

    #[test]
    fn decode_jwt_claims_standard_encoding_fallback() {
        use base64::Engine;
        let header = base64::engine::general_purpose::STANDARD.encode(r#"{"alg":"none"}"#);
        let payload =
            base64::engine::general_purpose::STANDARD.encode(r#"{"exp":99,"aud":"test"}"#);
        let jwt = format!("{header}.{payload}.sig");
        let claims = decode_jwt_claims(&jwt).unwrap();
        assert_eq!(claims.exp, Some(99));
        assert_eq!(claims.aud.as_deref(), Some("test"));
    }

    // ── TokenInfo::from_jwt ────────────────────────────────────────

    #[test]
    fn from_jwt_with_exp_claim() {
        let jwt = make_test_jwt(&json!({"exp": 1700000000}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert!(info.expires_at.is_some());
        assert_eq!(info.expires_at.unwrap().timestamp(), 1700000000,);
    }

    #[test]
    fn from_jwt_without_exp_claim() {
        let jwt = make_test_jwt(&json!({"aud": "something"}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::IdToken).unwrap();
        assert!(info.expires_at.is_none());
    }

    #[test]
    fn from_jwt_with_aud_claim() {
        let jwt = make_test_jwt(&json!({"aud": "my-audience"}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert_eq!(info.audience, "my-audience");
    }

    #[test]
    fn from_jwt_without_aud_claim() {
        let jwt = make_test_jwt(&json!({"exp": 123}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert_eq!(info.audience, "");
    }

    #[test]
    fn from_jwt_preserves_raw_and_token_type() {
        let jwt = make_test_jwt(&json!({}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::IdToken).unwrap();
        assert_eq!(info.raw, jwt);
        assert_eq!(info.token_type, TokenType::IdToken);
    }

    // ── TokenInfo::is_expired ──────────────────────────────────────

    #[test]
    fn is_expired_past() {
        let jwt = make_test_jwt(&json!({"exp": 0}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert!(info.is_expired());
    }

    #[test]
    fn is_expired_future() {
        // Use a timestamp far in the future
        let jwt = make_test_jwt(&json!({"exp": 4102444800i64})); // 2100-01-01
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert!(!info.is_expired());
    }

    #[test]
    fn is_expired_none_returns_true() {
        let jwt = make_test_jwt(&json!({}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        assert!(info.expires_at.is_none());
        assert!(info.is_expired());
    }

    // ── extract_tenant_id ──────────────────────────────────────────

    #[test]
    fn extract_tenant_id_present() {
        let jwt = make_test_jwt(&json!({"tid": "abc-123"}));
        let tid = extract_tenant_id(&jwt).unwrap();
        assert_eq!(tid, Some("abc-123".to_string()));
    }

    #[test]
    fn extract_tenant_id_absent() {
        let jwt = make_test_jwt(&json!({"exp": 1}));
        let tid = extract_tenant_id(&jwt).unwrap();
        assert!(tid.is_none());
    }

    // ── extract_username ───────────────────────────────────────────

    #[test]
    fn extract_username_upn() {
        let jwt = make_test_jwt(&json!({
            "upn": "user@example.com",
            "preferred_username": "other",
            "unique_name": "another"
        }));
        let name = extract_username(&jwt).unwrap();
        assert_eq!(name, "user@example.com");
    }

    #[test]
    fn extract_username_preferred_username_fallback() {
        let jwt = make_test_jwt(&json!({
            "preferred_username": "pref@example.com",
            "unique_name": "uniq"
        }));
        let name = extract_username(&jwt).unwrap();
        assert_eq!(name, "pref@example.com");
    }

    #[test]
    fn extract_username_unique_name_fallback() {
        let jwt = make_test_jwt(&json!({"unique_name": "uniq@example.com"}));
        let name = extract_username(&jwt).unwrap();
        assert_eq!(name, "uniq@example.com");
    }

    #[test]
    fn extract_username_no_claims_errors() {
        let jwt = make_test_jwt(&json!({"exp": 1}));
        let result = extract_username(&jwt);
        assert!(result.is_err());
        let err = result.unwrap_err().to_string();
        assert!(err.contains("no username claim"), "got: {err}");
    }

    // ── TokenSet::is_expired ───────────────────────────────────────

    fn make_token_info(exp: i64) -> TokenInfo {
        let jwt = make_test_jwt(&json!({"exp": exp, "aud": "test"}));
        TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap()
    }

    fn make_test_token_set(teams_exp: i64, skype_exp: i64, agg_exp: i64) -> TokenSet {
        TokenSet {
            teams: make_token_info(teams_exp),
            skype: make_token_info(skype_exp),
            chatsvcagg: make_token_info(agg_exp),
            outlook: None,
            profile: "test".to_string(),
            tenant_id: "tid-123".to_string(),
        }
    }

    #[test]
    fn token_set_all_valid() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(far_future, far_future, far_future);
        assert!(!ts.is_expired());
    }

    #[test]
    fn token_set_one_expired() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(0, far_future, far_future);
        assert!(ts.is_expired());
    }

    #[test]
    fn token_set_all_expired() {
        let ts = make_test_token_set(0, 0, 0);
        assert!(ts.is_expired());
    }

    // ── Debug redaction ────────────────────────────────────────────

    #[test]
    fn debug_token_info_redacts_raw() {
        let jwt = make_test_jwt(&json!({"exp": 1}));
        let info = TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap();
        let debug = format!("{info:?}");
        assert!(debug.contains("[REDACTED]"), "got: {debug}");
        assert!(!debug.contains(&jwt), "raw JWT leaked in debug output");
    }

    #[test]
    fn debug_token_set_redacts_all_raws() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(far_future, far_future, far_future);
        let debug = format!("{ts:?}");
        // Should contain REDACTED for each TokenInfo
        assert!(debug.contains("[REDACTED]"), "got: {debug}");
        // Raw tokens must not appear
        assert!(
            !debug.contains(&ts.teams.raw),
            "teams JWT leaked in debug output"
        );
        assert!(
            !debug.contains(&ts.skype.raw),
            "skype JWT leaked in debug output"
        );
        assert!(
            !debug.contains(&ts.chatsvcagg.raw),
            "chatsvcagg JWT leaked in debug output"
        );
    }

    // ── Outlook token (backwards compat) ──────────────────────────

    #[test]
    fn token_set_deserialize_without_outlook_field() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(far_future, far_future, far_future);
        let mut json_val: serde_json::Value = serde_json::to_value(&ts).unwrap();
        // Remove the outlook field to simulate old token files
        json_val.as_object_mut().unwrap().remove("outlook");
        let json = serde_json::to_string(&json_val).unwrap();
        let deserialized: TokenSet = serde_json::from_str(&json).unwrap();
        assert!(deserialized.outlook.is_none());
        assert_eq!(deserialized.profile, "test");
    }

    #[test]
    fn token_set_is_expired_ignores_outlook() {
        let far_future = 4102444800i64;
        let mut ts = make_test_token_set(far_future, far_future, far_future);
        // Outlook expired should not affect is_expired
        ts.outlook = Some(make_token_info(0));
        assert!(!ts.is_expired());
    }

    #[test]
    fn outlook_bearer_returns_error_when_none() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(far_future, far_future, far_future);
        assert!(ts.outlook_bearer().is_err());
    }

    #[test]
    fn outlook_bearer_returns_bearer_when_some() {
        let far_future = 4102444800i64;
        let mut ts = make_test_token_set(far_future, far_future, far_future);
        ts.outlook = Some(make_token_info(far_future));
        let bearer = ts.outlook_bearer().unwrap();
        assert!(bearer.starts_with("Bearer "));
    }
}
