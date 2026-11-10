mod api;
mod auth;
mod cli;
mod config;
mod error;
mod html;
mod models;
mod output;

use clap::Parser;
use tracing_subscriber::EnvFilter;

use cli::{Cli, Commands};
use output::OutputFormat;

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
    // Handle explicit `auth login`.
    if let Commands::Auth(ref auth_args) = cli.command {
        if let cli::auth::AuthCommand::Login { ref tenant } = auth_args.command {
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
    // refresh first (see auth::refresh — redeems the stored refresh token
    // against Azure AD, no UI involved). That's safe even for unattended
    // cron jobs run with --no-auto-login, so it always runs; only the
    // interactive webview fallback below respects that flag.
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

        // `Some(existing)` means tokens are stale and need attention;
        // `existing` itself is `None` when there's nothing stored at all
        // (vs. Some(ts) for a stored-but-expired set) — only the latter is
        // worth attempting a silent refresh against.
        let stale: Option<Option<auth::token::TokenSet>> = match auth::resolve_tokens(profile) {
            Ok(Some(ts)) => {
                let needs = ts.is_expired()
                    || (is_outlook_cmd && ts.outlook.as_ref().is_none_or(|t| t.is_expired()));
                needs.then_some(Some(ts))
            }
            Ok(None) => Some(None),
            Err(_) => Some(None),
        };

        if let Some(existing) = stale {
            let refreshed = existing.as_ref().and_then(|ts| {
                auth::refresh::refresh_token_set(ts)
                    .inspect_err(|e| tracing::debug!("silent token refresh failed: {e}"))
                    .ok()
            });

            match refreshed {
                Some(fresh) => {
                    if let Err(e) = auth::keyring::store_tokens(profile, &fresh) {
                        tracing::debug!("failed to persist refreshed tokens: {e}");
                    }
                }
                None if !cli.no_auto_login => {
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
                }
                None => {
                    // --no-auto-login and refresh failed or wasn't possible
                    // (no tokens / no refresh token stored): leave things
                    // as-is. run() surfaces the appropriate
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
