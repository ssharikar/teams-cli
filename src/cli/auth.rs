use clap::{Args, Subcommand};

use crate::auth;
use crate::auth::token::{extract_username, TokenSet};
use crate::error::Result;
use crate::output::{self, OutputFormat};

#[derive(Args)]
pub struct AuthArgs {
    #[command(subcommand)]
    pub command: AuthCommand,
}

#[derive(Subcommand)]
pub enum AuthCommand {
    /// Authenticate with Microsoft Teams and Outlook
    Login {
        /// Tenant ID or 'common'
        #[arg(long, default_value = "common")]
        tenant: String,
    },
    /// Show current authentication status
    Status,
    /// Clear stored credentials
    Logout {
        /// Clear all profiles
        #[arg(long)]
        all: bool,
    },
    /// Print raw access token for scripting
    Token {
        /// Which token to print: teams, skype, chatsvcagg, outlook
        #[arg(default_value = "skype")]
        token_type: String,
    },
    /// List all profiles with stored credentials
    List,
    /// Silently refresh tokens using the stored refresh token (no browser)
    Refresh,
}

pub async fn handle(args: &AuthArgs, profile: &str, format: OutputFormat) -> Result<()> {
    match &args.command {
        AuthCommand::Login { .. } => {
            // Webview login is handled on the main thread before the async
            // runtime starts. See main.rs.
            Ok(())
        }

        AuthCommand::Status => {
            match auth::resolve_tokens(profile)? {
                Some(ts) => {
                    let username =
                        extract_username(&ts.teams.raw).unwrap_or_else(|_| "unknown".into());

                    let status = serde_json::json!({
                        "logged_in": true,
                        "profile": ts.profile,
                        "username": username,
                        "tenant_id": ts.tenant_id,
                        "teams_token": {
                            "expired": ts.teams.is_expired(),
                            "expires_at": ts.teams.expires_at.map(|t| t.to_rfc3339()),
                        },
                        "skype_token": {
                            "expired": ts.skype.is_expired(),
                            "expires_at": ts.skype.expires_at.map(|t| t.to_rfc3339()),
                        },
                        "chatsvcagg_token": {
                            "expired": ts.chatsvcagg.is_expired(),
                            "expires_at": ts.chatsvcagg.expires_at.map(|t| t.to_rfc3339()),
                        },
                        "outlook_token": ts.outlook.as_ref().map(|t| serde_json::json!({
                            "expired": t.is_expired(),
                            "expires_at": t.expires_at.map(|e| e.to_rfc3339()),
                        })),
                        "refreshable": ts.is_refreshable(),
                    });
                    output::print_output(format, status, 0);
                }
                None => {
                    let status = serde_json::json!({
                        "logged_in": false,
                        "profile": profile,
                    });
                    output::print_output(format, status, 0);
                }
            }
            Ok(())
        }

        AuthCommand::Logout { all } => {
            if *all {
                let profiles = auth::keyring::list_profiles()?;
                for p in &profiles {
                    auth::keyring::delete_tokens(p)?;
                }
                eprintln!("Logged out of all profiles");
            } else {
                auth::keyring::delete_tokens(profile)?;
                eprintln!("Logged out of profile '{profile}'");
            }
            Ok(())
        }

        AuthCommand::Token { token_type } => {
            let ts = auth::resolve_tokens(profile)?
                .ok_or_else(|| crate::error::TeamsError::AuthError("not logged in".into()))?;

            let token = match token_type.as_str() {
                "teams" => &ts.teams.raw,
                "skype" => &ts.skype.raw,
                "chatsvcagg" => &ts.chatsvcagg.raw,
                "outlook" => match &ts.outlook {
                    Some(t) => &t.raw,
                    None => {
                        return Err(crate::error::TeamsError::AuthError(
                            "outlook token not available; run 'teams auth login' first".into(),
                        ));
                    }
                },
                other => {
                    return Err(crate::error::TeamsError::InvalidInput(format!(
                        "unknown token type: {other} (use teams, skype, chatsvcagg, or outlook)"
                    )));
                }
            };
            eprintln!("warning: raw authentication token written to stdout");
            println!("{token}");
            Ok(())
        }

        AuthCommand::List => {
            let profiles = auth::keyring::list_profiles()?;

            let display: Vec<serde_json::Value> = profiles
                .iter()
                .map(|p| match auth::resolve_tokens(p) {
                    Ok(Some(ts)) => {
                        let username =
                            extract_username(&ts.teams.raw).unwrap_or_else(|_| "unknown".into());
                        serde_json::json!({
                            "profile": p,
                            "username": username,
                            "expired": ts.teams.is_expired(),
                            "refreshable": ts.is_refreshable(),
                        })
                    }
                    _ => serde_json::json!({
                        "profile": p,
                        "username": null,
                        "expired": null,
                        "refreshable": null,
                    }),
                })
                .collect();

            output::print_output(format, display, 0);
            Ok(())
        }

        AuthCommand::Refresh => {
            let ts = auth::resolve_tokens(profile)?
                .ok_or_else(|| crate::error::TeamsError::AuthError("not logged in".into()))?;

            let fresh = auth::refresh::refresh_token_set(&ts)?;
            auth::keyring::store_tokens(profile, &fresh)?;

            let username = extract_username(&fresh.teams.raw).unwrap_or_else(|_| "unknown".into());
            let result = serde_json::json!({
                "status": "refreshed",
                "profile": fresh.profile,
                "username": username,
            });
            output::print_output(format, result, 0);
            Ok(())
        }
    }
}

pub fn print_login_success(token_set: &TokenSet, format: OutputFormat) {
    let username = extract_username(&token_set.teams.raw).unwrap_or_else(|_| "unknown".into());
    let result = serde_json::json!({
        "status": "authenticated",
        "profile": token_set.profile,
        "username": username,
        "tenant_id": token_set.tenant_id,
    });
    output::print_output(format, result, 0);
}
