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
/// stored refresh token (from `auth login --device-code`), attempts a
/// silent (no browser) refresh first — see `auth::refresh`. Returns error
/// if that fails or tokens are missing.
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

/// Checks the Outlook token on an already-resolved `TokenSet`. Split out
/// from `ensure_outlook_token` so the validity check can be unit tested
/// without going through keyring/env I/O.
fn check_outlook_token(tokens: TokenSet) -> Result<TokenSet> {
    match &tokens.outlook {
        Some(t) if !t.is_expired() => Ok(tokens),
        _ => Err(TeamsError::AuthError(
            "outlook token missing or expired; run 'teams auth login' to re-authenticate".into(),
        )),
    }
}

/// Ensure the Outlook token is present and valid. The webview login flow
/// acquires the Outlook token as its 4th phase, so this usually just
/// validates that it exists and isn't expired. Unlike `get_or_login`, this
/// does not require the Teams/Skype/ChatSvcAgg tokens to be valid:
/// mail/calendar commands only use the Outlook token. If it's stale, tries
/// a silent refresh (same as `get_or_login`) before giving up.
pub async fn ensure_outlook_token(profile: &str, _tenant: &str) -> Result<TokenSet> {
    let tokens = resolve_tokens(profile)?.ok_or_else(|| {
        TeamsError::AuthError("not logged in; run 'teams auth login'".into())
    })?;

    let outlook_ok = matches!(&tokens.outlook, Some(t) if !t.is_expired());
    if outlook_ok {
        return Ok(tokens);
    }

    if let Ok(fresh) = try_refresh(profile, &tokens) {
        return check_outlook_token(fresh);
    }
    check_outlook_token(tokens)
}

#[cfg(test)]
mod tests {
    use super::*;
    use token::{TokenInfo, TokenType};
    use chrono::{Duration, Utc};

    fn token(expires_in: Duration) -> TokenInfo {
        TokenInfo {
            raw: "test".into(),
            token_type: TokenType::AccessToken,
            expires_at: Some(Utc::now() + expires_in),
            audience: String::new(),
        }
    }

    fn token_set(outlook: Option<TokenInfo>) -> TokenSet {
        // Teams/skype/chatsvcagg are expired on purpose: check_outlook_token
        // must not care about them, only about the Outlook token.
        let expired = token(Duration::hours(-1));
        TokenSet {
            teams: expired.clone(),
            skype: expired.clone(),
            chatsvcagg: expired,
            outlook,
            refresh_token: None,
            profile: "default".into(),
            tenant_id: "common".into(),
        }
    }

    #[test]
    fn check_outlook_token_ok_despite_expired_teams_token() {
        let ts = token_set(Some(token(Duration::hours(1))));
        assert!(check_outlook_token(ts).is_ok());
    }

    #[test]
    fn check_outlook_token_fails_when_outlook_expired() {
        let ts = token_set(Some(token(Duration::hours(-1))));
        assert!(check_outlook_token(ts).is_err());
    }

    #[test]
    fn check_outlook_token_fails_when_outlook_missing() {
        let ts = token_set(None);
        assert!(check_outlook_token(ts).is_err());
    }
}
