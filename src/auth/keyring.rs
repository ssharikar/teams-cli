use std::fs;
use std::path::PathBuf;

use crate::error::{Result, TeamsError};

use super::token::TokenSet;

fn tokens_dir() -> PathBuf {
    dirs::config_dir()
        .or_else(|| dirs::home_dir().map(|h| h.join(".config")))
        .unwrap_or_else(|| PathBuf::from("/tmp"))
        .join("teams-cli")
        .join("tokens")
}

fn token_path(profile: &str) -> PathBuf {
    tokens_dir().join(format!("{profile}.json"))
}

fn validate_profile_name(profile: &str) -> Result<()> {
    if profile.is_empty()
        || profile.contains('/')
        || profile.contains('\\')
        || profile.contains("..")
        || profile.starts_with('.')
        || !profile
            .chars()
            .all(|c| c.is_alphanumeric() || c == '-' || c == '_')
    {
        return Err(TeamsError::InvalidInput(format!(
            "invalid profile name '{profile}': must contain only alphanumeric, dash, or underscore characters"
        )));
    }
    Ok(())
}

pub fn store_tokens(profile: &str, token_set: &TokenSet) -> Result<()> {
    validate_profile_name(profile)?;

    let dir = tokens_dir();

    #[cfg(unix)]
    {
        use std::fs::DirBuilder;
        use std::os::unix::fs::DirBuilderExt;
        DirBuilder::new()
            .mode(0o700)
            .recursive(true)
            .create(&dir)
            .map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    }

    #[cfg(not(unix))]
    {
        fs::create_dir_all(&dir).map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    }

    let json = serde_json::to_string_pretty(token_set)
        .map_err(|e| TeamsError::KeyringError(format!("failed to serialize tokens: {e}")))?;

    let path = token_path(profile);

    #[cfg(unix)]
    {
        use std::io::Write;
        use std::os::unix::fs::OpenOptionsExt;
        let mut file = fs::OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .mode(0o600)
            .open(&path)
            .map_err(|e| TeamsError::KeyringError(e.to_string()))?;
        file.write_all(json.as_bytes())
            .map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    }

    #[cfg(not(unix))]
    {
        fs::write(&path, &json).map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    }

    Ok(())
}

pub fn get_tokens(profile: &str) -> Result<Option<TokenSet>> {
    validate_profile_name(profile)?;

    let path = token_path(profile);
    if !path.exists() {
        return Ok(None);
    }

    let json = fs::read_to_string(&path).map_err(|e| TeamsError::KeyringError(e.to_string()))?;

    let token_set: TokenSet = serde_json::from_str(&json)
        .map_err(|e| TeamsError::KeyringError(format!("failed to parse tokens: {e}")))?;

    Ok(Some(token_set))
}

pub fn delete_tokens(profile: &str) -> Result<()> {
    validate_profile_name(profile)?;

    let path = token_path(profile);
    if path.exists() {
        fs::remove_file(&path).map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    }
    Ok(())
}

pub fn list_profiles() -> Result<Vec<String>> {
    let dir = tokens_dir();
    if !dir.exists() {
        return Ok(vec![]);
    }

    let mut profiles = vec![];
    let entries = fs::read_dir(&dir).map_err(|e| TeamsError::KeyringError(e.to_string()))?;
    for entry in entries.flatten() {
        if let Some(name) = entry.path().file_stem() {
            profiles.push(name.to_string_lossy().to_string());
        }
    }
    Ok(profiles)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::auth::token::{TokenInfo, TokenType};
    use serde_json::json;

    fn make_test_jwt(claims: &serde_json::Value) -> String {
        use base64::Engine;
        let header = base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(r#"{"alg":"none"}"#);
        let payload = base64::engine::general_purpose::URL_SAFE_NO_PAD
            .encode(serde_json::to_string(claims).unwrap());
        format!("{header}.{payload}.sig")
    }

    fn make_test_token_set() -> TokenSet {
        let far_future = 4102444800i64;
        let jwt_teams = make_test_jwt(&json!({"exp": far_future, "aud": "teams", "tid": "t-123"}));
        let jwt_skype = make_test_jwt(&json!({"exp": far_future, "aud": "skype"}));
        let jwt_agg = make_test_jwt(&json!({"exp": far_future, "aud": "chatsvcagg"}));

        TokenSet {
            teams: TokenInfo::from_jwt(&jwt_teams, TokenType::IdToken).unwrap(),
            skype: TokenInfo::from_jwt(&jwt_skype, TokenType::AccessToken).unwrap(),
            chatsvcagg: TokenInfo::from_jwt(&jwt_agg, TokenType::AccessToken).unwrap(),
            outlook: None,
            profile: "test-profile".to_string(),
            tenant_id: "t-123".to_string(),
        }
    }

    // ── validate_profile_name ──────────────────────────────────────

    #[test]
    fn valid_profile_names() {
        for name in &["default", "my-profile", "test_123", "A", "abc123"] {
            assert!(
                validate_profile_name(name).is_ok(),
                "expected Ok for '{name}'"
            );
        }
    }

    #[test]
    fn rejects_path_traversal() {
        for name in &["../evil", "../../etc"] {
            assert!(
                validate_profile_name(name).is_err(),
                "expected Err for '{name}'"
            );
        }
    }

    #[test]
    fn rejects_slashes() {
        assert!(validate_profile_name("foo/bar").is_err());
        assert!(validate_profile_name("foo\\bar").is_err());
    }

    #[test]
    fn rejects_dot_prefix() {
        assert!(validate_profile_name(".hidden").is_err());
    }

    #[test]
    fn rejects_empty_string() {
        assert!(validate_profile_name("").is_err());
    }

    #[test]
    fn rejects_special_chars() {
        for name in &["foo@bar", "hello world", "name!", "a+b", "x=y"] {
            assert!(
                validate_profile_name(name).is_err(),
                "expected Err for '{name}'"
            );
        }
    }

    #[test]
    fn error_message_mentions_profile_name() {
        let result = validate_profile_name("../evil");
        let err = result.unwrap_err().to_string();
        assert!(err.contains("../evil"), "got: {err}");
        assert!(err.contains("invalid profile name"), "got: {err}");
    }

    // ── TokenSet JSON round-trip ───────────────────────────────────

    #[test]
    fn token_set_json_round_trip() {
        let ts = make_test_token_set();
        let json = serde_json::to_string_pretty(&ts).unwrap();
        let deserialized: TokenSet = serde_json::from_str(&json).unwrap();

        assert_eq!(ts.profile, deserialized.profile);
        assert_eq!(ts.tenant_id, deserialized.tenant_id);
        assert_eq!(ts.teams.raw, deserialized.teams.raw);
        assert_eq!(ts.teams.audience, deserialized.teams.audience);
        assert_eq!(ts.teams.expires_at, deserialized.teams.expires_at);
        assert_eq!(ts.skype.raw, deserialized.skype.raw);
        assert_eq!(ts.skype.audience, deserialized.skype.audience);
        assert_eq!(ts.chatsvcagg.raw, deserialized.chatsvcagg.raw);
        assert_eq!(ts.chatsvcagg.audience, deserialized.chatsvcagg.audience);
    }

    #[test]
    fn token_set_json_round_trip_preserves_token_type() {
        let ts = make_test_token_set();
        let json = serde_json::to_string(&ts).unwrap();
        let deserialized: TokenSet = serde_json::from_str(&json).unwrap();

        assert_eq!(ts.teams.token_type, deserialized.teams.token_type);
        assert_eq!(ts.skype.token_type, deserialized.skype.token_type);
        assert_eq!(ts.chatsvcagg.token_type, deserialized.chatsvcagg.token_type);
    }

    // ── delete_tokens on non-existent profile ──────────────────────

    #[test]
    fn delete_tokens_nonexistent_does_not_error() {
        // This calls the real filesystem but the profile shouldn't exist.
        // validate_profile_name is the main thing being exercised; the
        // file just won't be there so the function returns Ok.
        let result = delete_tokens("nonexistent-profile-that-should-not-exist");
        assert!(result.is_ok());
    }
}
