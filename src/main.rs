mod api;
mod auth;
mod cli;
mod config;
mod error;
mod html;
mod models;
mod output;

use std::io::IsTerminal;

use clap::Parser;
use tracing_subscriber::EnvFilter;

use auth::token::TokenSet;
use cli::{Cli, Commands};
use output::OutputFormat;

// Whether auto-login should run, given the stored tokens (if any) and
// whether the command only needs the Outlook token (mail/calendar) rather
// than the Teams/Skype/ChatSvcAgg set.
fn needs_login_for(ts: Option<&TokenSet>, is_outlook_cmd: bool) -> bool {
    match ts {
        Some(ts) => {
            if is_outlook_cmd {
                ts.outlook.as_ref().is_none_or(|t| t.is_expired())
            } else {
                ts.is_expired()
            }
        }
        None => true,
    }
}

fn main() {
    let cli = Cli::parse();

    // Set up tracing
    let filter = match cli.verbose {
        0 => "warn",
        1 => "info",
        2 => "debug",
        _ => "trace",
    };
    tracing_subscriber::fmt()
        .with_env_filter(
            EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new(filter)),
        )
        .with_writer(std::io::stderr)
        .init();

    // DOC-6: Make --no-color functional via the NO_COLOR convention
    if cli.no_color || std::env::var("NO_COLOR").is_ok() {
        std::env::set_var("NO_COLOR", "1");
    }

    let format = match OutputFormat::detect(cli.output.as_deref()) {
        Ok(f) => f,
        Err(msg) => {
            eprintln!("error: {msg}");
            std::process::exit(2);
        }
    };

    // Webview login must run on the main thread (tao/wry requirement).
    // Handle explicit `auth login`, unless it's `--device-code`: that's a
    // plain HTTP polling loop with no GUI involved, so it's left to run
    // through the normal async dispatch in `run()` (cli::auth::handle)
    // instead of being intercepted here.
    if let Commands::Auth(ref auth_args) = cli.command {
        if let cli::auth::AuthCommand::Login {
            ref tenant,
            device_code: false,
        } = auth_args.command
        {
            match auth::webview::webview_login(tenant, &cli.profile) {
                Ok(token_set) => {
                    let username = auth::token::extract_username(&token_set.teams.raw)
                        .unwrap_or_else(|_| "unknown".into());
                    eprintln!("Authenticated as {username}");
                    cli::auth::print_login_success(&token_set, format);
                    return;
                }
                Err(e) => {
                    output::print_error(format, e.error_code(), &e.to_string(), 0);
                    std::process::exit(e.exit_code());
                }
            }
        }
    }

    // Keep tokens usable without necessarily popping a browser: if a
    // non-auth command's tokens are missing or expired, try a silent
    // refresh first (see auth::refresh — redeems a stored refresh token
    // against Azure AD, no UI involved, only present after `auth login
    // --device-code`). That's safe even for unattended agent invocations
    // run with --no-auto-login, so it always runs; only the interactive
    // webview fallback below respects that flag.
    if !matches!(
        cli.command,
        Commands::Auth(_) | Commands::Config(_) | Commands::Completions { .. }
    ) {
        let cfg = config::Config::load().ok();
        let profile = if cli.profile == "default" {
            cfg.as_ref()
                .map(|c| c.default.profile.as_str())
                .unwrap_or("default")
        } else {
            &cli.profile
        };
        let is_outlook_cmd = matches!(cli.command, Commands::Mail(_) | Commands::Calendar(_));
        let existing = auth::resolve_tokens(profile).ok().flatten();
        let needs_login = needs_login_for(existing.as_ref(), is_outlook_cmd);

        if needs_login {
            let refreshed = existing.as_ref().and_then(|ts| {
                auth::refresh::refresh_token_set(ts)
                    .inspect_err(|e| tracing::debug!("silent token refresh failed: {e}"))
                    .ok()
            });

            match refreshed {
                Some(fresh) => {
                    if let Err(e) = auth::keyring::store_tokens(profile, &fresh) {
                        tracing::debug!("failed to persist refreshed tokens: {e}");
                    } else {
                        tracing::debug!(
                            "silent refresh successful for profile '{}'",
                            fresh.profile
                        );
                    }
                }
                None if !cli.no_auto_login => {
                    // The interactive webview blocks until a human clicks
                    // through a sign-in page; with stdin not a tty there's
                    // nobody to do that, so it would hang indefinitely
                    // (e.g. a script, cron job, or agent invocation with
                    // no one watching the window) — skip it and fail fast
                    // instead.
                    if std::io::stdin().is_terminal() {
                        let tenant = cfg
                            .as_ref()
                            .map(|c| c.profile(profile).tenant_id.clone())
                            .unwrap_or_else(|| "common".to_string());
                        match auth::webview::webview_login(&tenant, profile) {
                            Ok(token_set) => {
                                tracing::debug!(
                                    "auto-login successful for profile '{}'",
                                    token_set.profile
                                );
                            }
                            Err(e) => {
                                eprintln!("auto-login failed: {e}");
                                std::process::exit(e.exit_code());
                            }
                        }
                    } else {
                        eprintln!(
                            "tokens expired or missing and stdin isn't a terminal, so the \
                             interactive login can't run here; run 'teams auth login' from \
                             an interactive terminal first, or set up 'teams auth login \
                             --device-code' once so future commands can refresh silently"
                        );
                    }
                }
                None => {
                    // --no-auto-login and refresh failed or wasn't
                    // possible (no tokens / no refresh token stored):
                    // leave things as-is. run() surfaces the appropriate
                    // TokenExpired/AuthError itself.
                }
            }
        }
    }

    let rt = tokio::runtime::Runtime::new().expect("failed to create tokio runtime");
    rt.block_on(async {
        if let Err(e) = run(cli, format).await {
            output::print_error(format, e.error_code(), &e.to_string(), 0);
            std::process::exit(e.exit_code());
        }
    });
}

async fn run(cli: Cli, format: OutputFormat) -> error::Result<()> {
    let cfg = config::Config::load()?;

    let profile = if cli.profile == "default" {
        &cfg.default.profile
    } else {
        &cli.profile
    };
    let _region = if cli.region == "emea" {
        &cfg.default.region
    } else {
        &cli.region
    };

    let network = config::NetworkConfig {
        timeout: cli.timeout.unwrap_or(cfg.network.timeout),
        max_retries: cli.retry.unwrap_or(cfg.network.max_retries),
        retry_backoff_base: cfg.network.retry_backoff_base,
    };

    match &cli.command {
        Commands::Auth(args) => cli::auth::handle(args, profile, format).await,
        Commands::Config(args) => cli::config_cmd::handle(args, format),
        Commands::Completions { shell } => {
            let mut cmd = <Cli as clap::CommandFactory>::command();
            clap_complete::generate(*shell, &mut cmd, "teams", &mut std::io::stdout());
            Ok(())
        }

        // Outlook commands use lazy token acquisition (no Teams authz exchange)
        Commands::Mail(args) => {
            let tenant = cfg.profile(profile).tenant_id.clone();
            let http = api::HttpClient::new(&network);
            let tokens = auth::ensure_outlook_token(profile, &tenant).await?;
            cli::mail::handle(args, &tokens, &http, format).await
        }
        Commands::Calendar(args) => {
            let tenant = cfg.profile(profile).tenant_id.clone();
            let http = api::HttpClient::new(&network);
            let tokens = auth::ensure_outlook_token(profile, &tenant).await?;
            cli::calendar::handle(args, &tokens, &http, format).await
        }
        // Presence uses skype_bearer for the forceavailability endpoint
        Commands::Presence(args) => {
            let tenant = cfg.profile(profile).tenant_id.clone();
            let tokens = auth::get_or_login(profile, &tenant, !cli.no_auto_login).await?;
            let http = api::HttpClient::new(&network);
            cli::presence::handle(args, &tokens, &http, format).await
        }
        // All other commands need auth + authz token exchange
        cmd => {
            let tenant = cfg.profile(profile).tenant_id.clone();
            let tokens = auth::get_or_login(profile, &tenant, !cli.no_auto_login).await?;
            let http = api::HttpClient::new(&network);

            // Exchange OAuth token for messaging skype token + discover region
            let authz = api::authz::exchange_token(&http, &tokens).await?;
            let chat_service_url = &authz.region_gtms.chat_service;
            let mt_url = &authz.region_gtms.middle_tier;
            let messaging_token = &authz.tokens.skype_token;

            match cmd {
                Commands::User(args) => {
                    cli::user::handle(args, &tokens, &http, mt_url, format).await
                }
                Commands::Team(args) => cli::team::handle(args, &tokens, &http, format).await,
                Commands::Channel(args) => cli::channel::handle(args, &tokens, &http, format).await,
                Commands::Chat(args) => {
                    // chat create needs messaging context for conversation creation
                    if let cli::chat::ChatCommand::Create { ref user } = args.command {
                        let chat_ctx = cli::chat::ChatContext {
                            tokens: &tokens,
                            messaging_token,
                            http: &http,
                            chat_service_url,
                            mt_url,
                        };
                        cli::chat::handle_create(user, &chat_ctx, format).await
                    } else {
                        cli::chat::handle(args, &tokens, &http, format).await
                    }
                }
                Commands::Message(args) => {
                    let msg_ctx = cli::message::MessageContext {
                        tokens: &tokens,
                        messaging_token,
                        http: &http,
                        chat_service_url,
                        ams_v2_url: &authz.region_gtms.ams_v2,
                        ams_url: &authz.region_gtms.ams,
                    };
                    cli::message::handle(args, &msg_ctx, format).await
                }
                Commands::Tenant(args) => {
                    cli::tenant::handle(args, &tokens, &http, mt_url, format).await
                }
                _ => unreachable!(),
            }
        }
    }
}

#[cfg(test)]
mod needs_login_tests {
    use super::needs_login_for;
    use crate::auth::token::{TokenInfo, TokenSet, TokenType};
    use chrono::{Duration, Utc};

    fn token(expires_in: Duration) -> TokenInfo {
        TokenInfo {
            raw: "test".into(),
            token_type: TokenType::AccessToken,
            expires_at: Some(Utc::now() + expires_in),
            audience: String::new(),
        }
    }

    fn token_set(teams_expired: bool, outlook: Option<bool>) -> TokenSet {
        let valid = Duration::hours(1);
        let expired = Duration::hours(-1);
        TokenSet {
            teams: token(if teams_expired { expired } else { valid }),
            skype: token(valid),
            chatsvcagg: token(valid),
            outlook: outlook.map(|outlook_expired| {
                token(if outlook_expired { expired } else { valid })
            }),
            refresh_token: None,
            profile: "default".into(),
            tenant_id: "common".into(),
        }
    }

    #[test]
    fn no_tokens_needs_login() {
        assert!(needs_login_for(None, false));
        assert!(needs_login_for(None, true));
    }

    #[test]
    fn outlook_cmd_ignores_expired_teams_token() {
        // Teams token expired, Outlook token still valid: mail/calendar
        // should NOT be gated on the Teams token (the regression this
        // test guards against).
        let ts = token_set(true, Some(false));
        assert!(!needs_login_for(Some(&ts), true));
    }

    #[test]
    fn outlook_cmd_needs_login_when_outlook_expired() {
        let ts = token_set(false, Some(true));
        assert!(needs_login_for(Some(&ts), true));
    }

    #[test]
    fn outlook_cmd_needs_login_when_outlook_missing() {
        let ts = token_set(false, None);
        assert!(needs_login_for(Some(&ts), true));
    }

    #[test]
    fn teams_cmd_needs_login_when_teams_expired() {
        let ts = token_set(true, Some(false));
        assert!(needs_login_for(Some(&ts), false));
    }

    #[test]
    fn teams_cmd_ignores_outlook_expiry() {
        // A non-mail/calendar command doesn't use Outlook, so an expired
        // (or missing) Outlook token shouldn't force a login.
        let ts = token_set(false, Some(true));
        assert!(!needs_login_for(Some(&ts), false));

        let ts_no_outlook = token_set(false, None);
        assert!(!needs_login_for(Some(&ts_no_outlook), false));
    }

    #[test]
    fn teams_cmd_all_valid_no_login_needed() {
        let ts = token_set(false, Some(false));
        assert!(!needs_login_for(Some(&ts), false));
    }
}
