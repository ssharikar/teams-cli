pub mod keyring;
pub mod refresh;
pub mod token;
pub mod webview;

use crate::error::{Result, TeamsError};
use token::TokenSet;

/// Resolve tokens for the given profile.
/// Priority: env vars > keyring (if valid) > keyring (expired)
pub fn resolve_tokens(profile: &str) -> Result<Option<TokenSet>> {
    // Check environment variables first
    if let (Ok(teams), Ok(skype), Ok(chatsvcagg)) = (
        std::env::var("TEAMS_CLI_TEAMS_TOKEN"),
        std::env::var("TEAMS_CLI_SKYPE_TOKEN"),
        std::env::var("TEAMS_CLI_CHATSVCAGG_TOKEN"),
    ) {
        let outlook = std::env::var("TEAMS_CLI_OUTLOOK_TOKEN")
            .ok()
            .map(|t| token::TokenInfo::from_jwt(&t, token::TokenType::AccessToken))
            .transpose()?;

        let token_set = TokenSet {
            teams: token::TokenInfo::from_jwt(&teams, token::TokenType::IdToken)?,
            skype: token::TokenInfo::from_jwt(&skype, token::TokenType::AccessToken)?,
            chatsvcagg: token::TokenInfo::from_jwt(&chatsvcagg, token::TokenType::AccessToken)?,
            outlook,
            refresh_token: None,
            profile: profile.to_string(),
            tenant_id: token::extract_tenant_id(&teams)?.unwrap_or_else(|| "common".to_string()),
        };
        return Ok(Some(token_set));
    }

    // Check keyring (this is where webview login stores tokens)
    keyring::get_tokens(profile)
}

/// Get tokens if they exist and are valid. If they're expired but have a
/// stored refresh token, attempts a silent (no browser) refresh first —
/// see `auth::refresh`. Returns error if that fails or tokens are missing.
pub async fn get_or_login(profile: &str, _tenant: &str, _auto_login: bool) -> Result<TokenSet> {
    match resolve_tokens(profile)? {
        Some(ts) if !ts.is_expired() => Ok(ts),
        Some(ts) => try_refresh(profile, &ts).map_err(|_| TeamsError::TokenExpired),
        None => Err(TeamsError::AuthError(
            "not logged in; run 'teams auth login'".into(),
        )),
    }
}

/// Attempts a silent refresh and persists the result on success.
fn try_refresh(profile: &str, ts: &TokenSet) -> Result<TokenSet> {
    let fresh = refresh::refresh_token_set(ts).inspect_err(|e| {
        tracing::debug!("silent token refresh failed: {e}");
    })?;
    keyring::store_tokens(profile, &fresh)?;
    Ok(fresh)
}

/// Ensure the Outlook token is present and valid. The webview login flow
/// acquires the Outlook token as its 4th phase, so this usually just
/// validates that it exists and isn't expired — but if it's stale, it
/// tries a silent refresh (same as `get_or_login`) before giving up.
pub async fn ensure_outlook_token(profile: &str, tenant: &str) -> Result<TokenSet> {
    let tokens = get_or_login(profile, tenant, true).await?;

    if let Some(t) = &tokens.outlook {
        if !t.is_expired() {
            return Ok(tokens);
        }
    }

    let outlook_missing_err = || {
        TeamsError::AuthError(
            "outlook token missing or expired; run 'teams auth login' to re-authenticate".into(),
        )
    };

    match try_refresh(profile, &tokens) {
        Ok(fresh) => match &fresh.outlook {
            Some(t) if !t.is_expired() => Ok(fresh),
            _ => Err(outlook_missing_err()),
        },
        Err(_) => Err(outlook_missing_err()),
    }
}
