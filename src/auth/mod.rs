pub mod keyring;
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
            profile: profile.to_string(),
            tenant_id: token::extract_tenant_id(&teams)?.unwrap_or_else(|| "common".to_string()),
        };
        return Ok(Some(token_set));
    }

    // Check keyring (this is where webview login stores tokens)
    keyring::get_tokens(profile)
}

/// Get tokens if they exist and are valid. Returns error if expired or missing.
pub async fn get_or_login(profile: &str, _tenant: &str, _auto_login: bool) -> Result<TokenSet> {
    match resolve_tokens(profile)? {
        Some(ts) if !ts.is_expired() => Ok(ts),
        Some(_) => Err(TeamsError::TokenExpired),
        None => Err(TeamsError::AuthError(
            "not logged in; run 'teams auth login'".into(),
        )),
    }
}

/// Ensure the Outlook token is present and valid. The webview login flow
/// acquires the Outlook token as its 4th phase, so this just validates
/// that it exists and isn't expired.
pub async fn ensure_outlook_token(profile: &str, tenant: &str) -> Result<TokenSet> {
    let tokens = get_or_login(profile, tenant, true).await?;

    match &tokens.outlook {
        Some(t) if !t.is_expired() => Ok(tokens),
        _ => Err(TeamsError::AuthError(
            "outlook token missing or expired; run 'teams auth login' to re-authenticate".into(),
        )),
    }
}
