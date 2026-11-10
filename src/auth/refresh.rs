//! Silent token refresh via the Azure AD v1.0 token endpoint.
//!
//! The webview login flow (`webview.rs`) authenticates via OAuth2
//! implicit/hybrid grant to get the four access tokens this CLI needs, and
//! separately exchanges a one-time authorization code (also obtained during
//! that flow) for a refresh token. That refresh token is tied to the app's
//! client ID, not to a single resource — the same v1.0 token endpoint lets
//! it be redeemed with `grant_type=refresh_token` for *any* resource the
//! user already has access to (Skype, ChatSvcAgg, Outlook, ...), with no
//! further browser interaction. This module does that redemption, so a
//! command doesn't need a fresh interactive `auth login` every ~1 hour.
//!
//! All requests here are blocking (`reqwest::blocking`). This is used both
//! pre-`tokio` (the webview login/auto-login gate in `main.rs`, which must
//! run on the main thread before the async runtime starts) and from inside
//! `async fn`s elsewhere in `auth/`; the latter briefly blocks a tokio
//! worker thread, which is acceptable for a single in-flight CLI command.

use serde::Deserialize;

use super::token::{
    TokenInfo, TokenSet, TokenType, CHATSVCAGG_RESOURCE, OUTLOOK_RESOURCE, REDIRECT_URI,
    SKYPE_RESOURCE, TEAMS_APP_ID,
};
use crate::error::{Result, TeamsError};

const TOKEN_ENDPOINT_TEMPLATE: &str = "https://login.microsoftonline.com/{tenant}/oauth2/token";

fn token_endpoint(tenant: &str) -> String {
    TOKEN_ENDPOINT_TEMPLATE.replace("{tenant}", tenant)
}

#[derive(Debug, Deserialize, Default)]
struct TokenEndpointResponse {
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

    if !resp.status().is_success() {
        let body = resp.text().unwrap_or_default();
        return Err(TeamsError::AuthError(format!(
            "token request rejected: {body}"
        )));
    }

    resp.json()
        .map_err(|e| TeamsError::AuthError(format!("failed to parse token response: {e}")))
}

/// One-time exchange of the authorization code obtained during webview
/// login for a refresh token. The access token in the response is
/// discarded — the webview flow already captured one from the URL fragment
/// for the same resource; this call exists solely to get `refresh_token`.
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
    post_token_request(tenant, &form)
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

    // Outlook is lazily acquired by the webview login's 4th phase, so it
    // should normally already be present, but refresh it unconditionally —
    // a profile that only ever ran Teams commands could still lack it.
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
    }

    #[test]
    fn token_endpoint_response_deserializes_all_fields() {
        let json = r#"{"access_token": "a", "id_token": "i", "refresh_token": "r"}"#;
        let resp: TokenEndpointResponse = serde_json::from_str(json).unwrap();
        assert_eq!(resp.access_token.as_deref(), Some("a"));
        assert_eq!(resp.id_token.as_deref(), Some("i"));
        assert_eq!(resp.refresh_token.as_deref(), Some("r"));
    }
}
