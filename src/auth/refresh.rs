//! Silent token refresh and device-code bootstrap, both via the Azure AD
//! v1.0 token endpoint.
//!
//! The webview login flow (`webview.rs`) authenticates via OAuth2 implicit
//! grant, which never returns a refresh token — so historically every
//! command, once the ~1hr access tokens expired, had to reopen an
//! interactive browser window to sign in again.
//!
//! This module provides an alternative, one-time bootstrap
//! (`device_code_login`) using the OAuth2 Device Authorization Grant: the
//! user completes sign-in on *any* browser (their phone's, even) by
//! visiting a URL and typing a short code — no webview or GUI session is
//! needed on this machine at all. That exchange returns a real refresh
//! token tied to the app's client ID, not to a single resource, so the
//! same one can be redeemed via `grant_type=refresh_token` for *any*
//! resource the user has access to (Skype, ChatSvcAgg, Outlook, the Teams
//! id_token, ...) with no further browser interaction — see
//! `refresh_token_set`, used by `get_or_login`/`ensure_outlook_token` and
//! the auto-login gate in `main.rs` so a command doesn't need a fresh
//! interactive `auth login` every ~1 hour.
//!
//! All requests here are blocking (`reqwest::blocking`). This is used both
//! pre-`tokio` (the auto-login gate in `main.rs`, which must run on the
//! main thread before the async runtime starts) and from inside `async
//! fn`s elsewhere in `auth/`; the latter briefly blocks a tokio worker
//! thread, which is acceptable for a single in-flight CLI command.

use std::thread::sleep;
use std::time::{Duration, Instant};

use serde::Deserialize;

use super::token::{
    extract_tenant_id, TokenInfo, TokenSet, TokenType, CHATSVCAGG_RESOURCE, OUTLOOK_RESOURCE,
    REDIRECT_URI, SKYPE_RESOURCE, TEAMS_APP_ID,
};
use crate::error::{Result, TeamsError};

const TOKEN_ENDPOINT_TEMPLATE: &str = "https://login.microsoftonline.com/{tenant}/oauth2/token";
const DEVICE_CODE_ENDPOINT_TEMPLATE: &str =
    "https://login.microsoftonline.com/{tenant}/oauth2/devicecode?api-version=1.0";

fn token_endpoint(tenant: &str) -> String {
    TOKEN_ENDPOINT_TEMPLATE.replace("{tenant}", tenant)
}

fn device_code_endpoint(tenant: &str) -> String {
    DEVICE_CODE_ENDPOINT_TEMPLATE.replace("{tenant}", tenant)
}

#[derive(Debug, Deserialize, Default)]
struct TokenEndpointResponse {
    #[serde(default)]
    error: Option<String>,
    #[serde(default)]
    access_token: Option<String>,
    #[serde(default)]
    id_token: Option<String>,
    #[serde(default)]
    refresh_token: Option<String>,
}

fn post_token_request(tenant: &str, form: &[(&str, &str)]) -> Result<TokenEndpointResponse> {
    let client = reqwest::blocking::Client::new();
    let resp = client
        .post(token_endpoint(tenant))
        .form(form)
        .send()
        .map_err(|e| TeamsError::AuthError(format!("token request failed: {e}")))?;

    let status_ok = resp.status().is_success();
    let parsed: TokenEndpointResponse = resp
        .json()
        .map_err(|e| TeamsError::AuthError(format!("failed to parse token response: {e}")))?;

    if !status_ok && parsed.error.is_none() {
        return Err(TeamsError::AuthError("token request rejected".into()));
    }

    Ok(parsed)
}

/// One-time exchange of the authorization code obtained during the
/// webview's hybrid-response login (`webview.rs`) for a refresh token. The
/// access token in the response is discarded — the webview flow already
/// captured one from the URL fragment for the same resource; this call
/// exists solely to get `refresh_token`.
pub fn exchange_code_for_refresh_token(
    tenant: &str,
    code: &str,
    resource: &str,
) -> Result<Option<String>> {
    let form = [
        ("grant_type", "authorization_code"),
        ("client_id", TEAMS_APP_ID),
        ("code", code),
        ("redirect_uri", REDIRECT_URI),
        ("resource", resource),
    ];

    let resp = post_token_request(tenant, &form)?;
    if let Some(err) = &resp.error {
        return Err(TeamsError::AuthError(format!(
            "authorization code exchange rejected: {err}"
        )));
    }
    Ok(resp.refresh_token)
}

fn redeem_refresh_token(
    tenant: &str,
    refresh_token: &str,
    resource: Option<&str>,
) -> Result<TokenEndpointResponse> {
    let mut form: Vec<(&str, &str)> = vec![
        ("grant_type", "refresh_token"),
        ("client_id", TEAMS_APP_ID),
        ("refresh_token", refresh_token),
    ];
    if let Some(r) = resource {
        form.push(("resource", r));
    }
    let resp = post_token_request(tenant, &form)?;
    if let Some(err) = &resp.error {
        return Err(TeamsError::AuthError(format!("refresh rejected: {err}")));
    }
    Ok(resp)
}

/// Silently refreshes every token in `ts` using its stored refresh token,
/// returning a brand-new `TokenSet`. AAD may rotate the refresh token on
/// each redemption, so the latest one returned is chained into the next
/// call and persisted in the result.
///
/// Returns an error (no network calls attempted beyond what already ran) if
/// there's no refresh token to use, or AAD rejects it outright (e.g.
/// revoked, or the user changed their password) — callers should fall back
/// to an interactive `webview::webview_login` in that case.
pub fn refresh_token_set(ts: &TokenSet) -> Result<TokenSet> {
    let mut refresh_token = ts
        .refresh_token
        .clone()
        .ok_or_else(|| TeamsError::AuthError("no refresh token stored for this profile".into()))?;

    let skype_resp = redeem_refresh_token(&ts.tenant_id, &refresh_token, Some(SKYPE_RESOURCE))?;
    if let Some(rt) = skype_resp.refresh_token {
        refresh_token = rt;
    }
    let skype_access = skype_resp.access_token.ok_or_else(|| {
        TeamsError::AuthError("refresh response missing skype access_token".into())
    })?;

    let agg_resp = redeem_refresh_token(&ts.tenant_id, &refresh_token, Some(CHATSVCAGG_RESOURCE))?;
    if let Some(rt) = agg_resp.refresh_token {
        refresh_token = rt;
    }
    let agg_access = agg_resp.access_token.ok_or_else(|| {
        TeamsError::AuthError("refresh response missing chatsvcagg access_token".into())
    })?;

    // Best-effort: the "teams" field is a locally-decoded id_token used only
    // to display/extract the signed-in username — it's never sent to a live
    // API as a credential. If AAD doesn't hand back a fresh id_token on a
    // bare refresh (resource omitted), keep the stale one rather than
    // failing the whole refresh over it.
    let teams_resp = redeem_refresh_token(&ts.tenant_id, &refresh_token, None);
    let teams_raw = match &teams_resp {
        Ok(resp) => {
            if let Some(rt) = &resp.refresh_token {
                refresh_token = rt.clone();
            }
            resp.id_token
                .clone()
                .unwrap_or_else(|| ts.teams.raw.clone())
        }
        Err(_) => ts.teams.raw.clone(),
    };

    // Outlook is lazily acquired on first use, so it should normally
    // already be present, but refresh it unconditionally — a profile that
    // only ever ran Teams commands could still lack it.
    let outlook_resp = redeem_refresh_token(&ts.tenant_id, &refresh_token, Some(OUTLOOK_RESOURCE));
    let outlook = match outlook_resp {
        Ok(resp) => {
            if let Some(rt) = resp.refresh_token {
                refresh_token = rt;
            }
            resp.access_token
                .map(|raw| TokenInfo::from_jwt(&raw, TokenType::AccessToken))
                .transpose()?
        }
        Err(_) => None,
    };

    Ok(TokenSet {
        teams: TokenInfo::from_jwt(&teams_raw, TokenType::IdToken)?,
        skype: TokenInfo::from_jwt(&skype_access, TokenType::AccessToken)?,
        chatsvcagg: TokenInfo::from_jwt(&agg_access, TokenType::AccessToken)?,
        outlook,
        refresh_token: Some(refresh_token),
        profile: ts.profile.clone(),
        tenant_id: ts.tenant_id.clone(),
    })
}

#[derive(Debug, Deserialize)]
struct DeviceCodeResponse {
    device_code: String,
    user_code: String,
    verification_url: String,
    expires_in: String,
    interval: String,
    message: String,
}

fn request_device_code(tenant: &str, resource: &str) -> Result<DeviceCodeResponse> {
    let client = reqwest::blocking::Client::new();
    let resp = client
        .post(device_code_endpoint(tenant))
        .form(&[("client_id", TEAMS_APP_ID), ("resource", resource)])
        .send()
        .map_err(|e| TeamsError::AuthError(format!("device code request failed: {e}")))?;

    if !resp.status().is_success() {
        let body = resp.text().unwrap_or_default();
        return Err(TeamsError::AuthError(format!(
            "device code request rejected: {body}"
        )));
    }

    resp.json().map_err(|e| {
        TeamsError::AuthError(format!("failed to parse device code response: {e}"))
    })
}

/// Polls the token endpoint with `grant_type=device_code` until the user
/// completes sign-in (or the code expires/is declined). Blocking, by
/// design — this is only ever called from the dedicated `auth login
/// --device-code` command, never from the auto-login gate.
fn poll_device_code(
    tenant: &str,
    device_code: &str,
    interval_secs: u64,
    expires_in_secs: u64,
) -> Result<TokenEndpointResponse> {
    let deadline = Instant::now() + Duration::from_secs(expires_in_secs);
    let mut interval = Duration::from_secs(interval_secs.max(1));

    loop {
        sleep(interval);

        let resp = post_token_request(
            tenant,
            &[
                ("grant_type", "device_code"),
                ("client_id", TEAMS_APP_ID),
                ("code", device_code),
            ],
        )?;

        if resp.access_token.is_some() {
            return Ok(resp);
        }

        match resp.error.as_deref() {
            Some("authorization_pending") => {
                if Instant::now() >= deadline {
                    return Err(TeamsError::AuthError(
                        "device code expired before sign-in completed".into(),
                    ));
                }
            }
            Some("slow_down") => {
                interval += Duration::from_secs(5);
            }
            Some(other) => {
                return Err(TeamsError::AuthError(format!(
                    "device code sign-in failed: {other}"
                )));
            }
            None => {
                return Err(TeamsError::AuthError(
                    "device code poll returned no access token and no error".into(),
                ));
            }
        }
    }
}

/// One-time bootstrap: requests a device code, prints the sign-in URL/code
/// for the user to complete on any browser (phone included — no webview or
/// GUI session needed on this machine), polls until they do, then chains
/// the resulting refresh token across the other resources so a full
/// four-token `TokenSet` comes back, just like an interactive login would.
/// Persists the result to the keyring before returning. From then on,
/// `refresh_token_set` can redeem it silently, with no browser or tty.
pub fn device_code_login(tenant: &str, profile: &str) -> Result<TokenSet> {
    let dc = request_device_code(tenant, SKYPE_RESOURCE)?;
    eprintln!("{}", dc.message);

    let interval: u64 = dc.interval.parse().unwrap_or(5);
    let expires_in: u64 = dc.expires_in.parse().unwrap_or(900);

    let skype_resp = poll_device_code(tenant, &dc.device_code, interval, expires_in)?;
    let skype_access = skype_resp.access_token.ok_or_else(|| {
        TeamsError::AuthError("device code exchange missing skype access_token".into())
    })?;
    let mut refresh_token = skype_resp.refresh_token.ok_or_else(|| {
        TeamsError::AuthError("device code exchange did not return a refresh token".into())
    })?;

    let agg_resp = redeem_refresh_token(tenant, &refresh_token, Some(CHATSVCAGG_RESOURCE))?;
    if let Some(rt) = agg_resp.refresh_token {
        refresh_token = rt;
    }
    let agg_access = agg_resp.access_token.ok_or_else(|| {
        TeamsError::AuthError("refresh response missing chatsvcagg access_token".into())
    })?;

    let teams_resp = redeem_refresh_token(tenant, &refresh_token, None)?;
    if let Some(rt) = teams_resp.refresh_token {
        refresh_token = rt;
    }
    let teams_raw = teams_resp.id_token.ok_or_else(|| {
        TeamsError::AuthError("device code bootstrap did not return a teams id_token".into())
    })?;

    let outlook_resp = redeem_refresh_token(tenant, &refresh_token, Some(OUTLOOK_RESOURCE));
    let outlook = match outlook_resp {
        Ok(resp) => {
            if let Some(rt) = resp.refresh_token {
                refresh_token = rt;
            }
            resp.access_token
                .map(|raw| TokenInfo::from_jwt(&raw, TokenType::AccessToken))
                .transpose()?
        }
        Err(_) => None,
    };

    let tenant_id = extract_tenant_id(&teams_raw)?.unwrap_or_else(|| tenant.to_string());

    let ts = TokenSet {
        teams: TokenInfo::from_jwt(&teams_raw, TokenType::IdToken)?,
        skype: TokenInfo::from_jwt(&skype_access, TokenType::AccessToken)?,
        chatsvcagg: TokenInfo::from_jwt(&agg_access, TokenType::AccessToken)?,
        outlook,
        refresh_token: Some(refresh_token),
        profile: profile.to_string(),
        tenant_id,
    };

    super::keyring::store_tokens(profile, &ts)?;
    Ok(ts)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn token_endpoint_interpolates_tenant() {
        assert_eq!(
            token_endpoint("common"),
            "https://login.microsoftonline.com/common/oauth2/token"
        );
        assert_eq!(
            token_endpoint("my-tenant-id"),
            "https://login.microsoftonline.com/my-tenant-id/oauth2/token"
        );
    }

    #[test]
    fn device_code_endpoint_interpolates_tenant() {
        assert_eq!(
            device_code_endpoint("my-tenant-id"),
            "https://login.microsoftonline.com/my-tenant-id/oauth2/devicecode?api-version=1.0"
        );
    }

    #[test]
    fn refresh_token_set_errors_without_refresh_token() {
        let far_future = 4102444800i64;
        let ts = make_test_token_set(far_future);
        let err = refresh_token_set(&ts).unwrap_err();
        assert!(err.to_string().contains("no refresh token"));
    }

    fn make_test_token_set(exp: i64) -> TokenSet {
        use crate::auth::token::TokenInfo;
        use base64::Engine;

        fn make_jwt(claims: &serde_json::Value) -> String {
            let header =
                base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(r#"{"alg":"none"}"#);
            let payload = base64::engine::general_purpose::URL_SAFE_NO_PAD
                .encode(serde_json::to_string(claims).unwrap());
            format!("{header}.{payload}.sig")
        }

        let jwt = make_jwt(&serde_json::json!({"exp": exp, "aud": "test"}));
        TokenSet {
            teams: TokenInfo::from_jwt(&jwt, TokenType::IdToken).unwrap(),
            skype: TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap(),
            chatsvcagg: TokenInfo::from_jwt(&jwt, TokenType::AccessToken).unwrap(),
            outlook: None,
            refresh_token: None,
            profile: "test".to_string(),
            tenant_id: "common".to_string(),
        }
    }

    #[test]
    fn token_endpoint_response_deserializes_partial_fields() {
        let json = r#"{"access_token": "abc", "token_type": "Bearer", "expires_in": "3600"}"#;
        let resp: TokenEndpointResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.access_token.as_deref(), Some("abc"));
        assert!(resp.id_token.is_none());
        assert!(resp.refresh_token.is_none());
        assert!(resp.error.is_none());
    }

    #[test]
    fn token_endpoint_response_deserializes_all_fields() {
        let json = r#"{"access_token": "a", "id_token": "i", "refresh_token": "r"}"#;
        let resp: TokenEndpointResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.access_token.as_deref(), Some("a"));
        assert_eq!(resp.id_token.as_deref(), Some("i"));
        assert_eq!(resp.refresh_token.as_deref(), Some("r"));
    }

    #[test]
    fn token_endpoint_response_deserializes_error_field() {
        let json = r#"{"error": "authorization_pending"}"#;
        let resp: TokenEndpointResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.error.as_deref(), Some("authorization_pending"));
        assert!(resp.access_token.is_none());
    }

    #[test]
    fn device_code_response_deserializes_real_shape() {
        // Real response shape from Azure AD v1.0 devicecode endpoint:
        // expires_in/interval come back as JSON strings, not numbers.
        let json = r#"{
            "user_code": "ABCDE123",
            "device_code": "long-opaque-code",
            "verification_url": "https://login.microsoft.com/device",
            "expires_in": "900",
            "interval": "5",
            "message": "To sign in, use a web browser..."
        }"#;
        let resp: DeviceCodeResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.user_code, "ABCDE123");
        assert_eq!(resp.expires_in, "900");
        assert_eq!(resp.interval, "5");
    }
}
