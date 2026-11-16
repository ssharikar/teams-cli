# AGENTS.md

## Project Overview

teams-cli is a Rust CLI for Microsoft Teams and Outlook. Teams commands use
internal Skype/CSA APIs (not Microsoft Graph). Outlook commands (mail, calendar)
use the Outlook REST API v2.0 at `outlook.office365.com/api/v2.0/`.

## Build and Test

```sh
cargo build
cargo test --all-targets
cargo fmt -- --check
cargo clippy --all-targets -- -D warnings
```

## Architecture

```
src/
  main.rs           -- Entry point. Webview login runs before tokio starts (main thread requirement).
                       Auto-login before commands (no retry on failure). NO_COLOR handling.
  error.rs          -- TeamsError enum with deterministic exit codes (0-10).
  config.rs         -- TOML config at ~/.config/teams-cli/config.toml. Sentinel-based defaults.
  cli/              -- clap command definitions and handlers
    mod.rs          -- Cli struct, Commands enum, global flags (--no-auto-login, --no-color, etc.)
    auth.rs         -- login (webview), status, list, logout, token
    user.rs         -- me, get, search
    team.rs         -- list, get
    channel.rs      -- list, get, pinned
    chat.rs         -- list (--all for hidden), get, create
    message.rs      -- list, send (--body or --stdin), get, edit, delete, react, unreact;
                       conversation_id also accepts "#channel-name" or a chat
                       title/participant name, resolved via CSA conversations
    mail.rs         -- list, read, send, draft, send-draft, delete, download,
                       reply, forward, search (Outlook email)
    calendar.rs     -- list, get, create, update, cancel, respond (Outlook calendar)
    tenant.rs       -- list, domains
    config_cmd.rs   -- init, show, set, path
  auth/
    mod.rs          -- Token resolution: env vars > file store
    token.rs        -- TokenSet, TokenInfo, JWT decode, constants (incl. OUTLOOK_RESOURCE)
    webview.rs      -- tao/wry webview for 4-token OAuth2 implicit flow (returns Result<TokenSet>)
    keyring.rs      -- File-based token storage at ~/.config/teams-cli/tokens/<profile>.json (0600/0700)
  api/
    mod.rs          -- HttpClient with retry/backoff, percent-encoded URL params, domain-validated authz
    authz.rs        -- Token exchange: POST authsvc/v1.0/authz -> skypeToken + regionGtms
    csa.rs          -- Chat Service Aggregator: teams, channels, chats
    messages.rs     -- Messages API: read/write messages (uses authz skypeToken)
    mt.rs           -- MiddleTier: user profiles, tenants (uses authz-discovered MT URL)
    outlook.rs      -- OutlookClient: email + calendar via Outlook REST API v2.0
  models/           -- Serde structs for API responses
    mod.rs          -- Re-exports
    outlook.rs      -- OutlookMessage, OutlookEvent, SendMailRequest, etc.
    chat.rs         -- Chat thread models
    conversation.rs -- Conversation/thread models
    message.rs      -- Message models
    team.rs         -- Team/channel models
    user.rs         -- User/tenant models
  output/           -- JSON envelope, table, plain text formatters
    mod.rs          -- OutputFormat enum (detect errors on unrecognized format strings)
    json.rs         -- JSON envelope formatter
    table.rs        -- Human-readable table formatter
    plain.rs        -- Tab-separated plain text formatter
```

## Authentication Flow

1. **Webview login** (`teams auth login`): Opens native webview via tao/wry,
   navigates to `login.microsoftonline.com/oauth2/authorize` four times to
   capture Teams id_token, Skype access_token, ChatSvcAgg access_token, and
   Outlook access_token via OAuth2 implicit flow. Tokens stored to file
   (0600 permissions, 0700 directory).

2. **Authz exchange** (on every API command): POST to
   `teams.microsoft.com/api/authsvc/v1.0/authz` with the OAuth Skype token.
   Returns a messaging-capable `skypeToken` (24h TTL) and `regionGtms` with
   correct regional base URLs. Authz URLs are validated against Microsoft/Skype
   domains.

3. **Auto-login**: If tokens are missing or expired, the CLI attempts
   webview authentication before running the command, unless
   `--no-auto-login` is set OR stdin isn't a terminal
   (`std::io::IsTerminal`) - a non-interactive invocation has nobody to
   complete the sign-in window, so it fails fast instead of hanging on an
   unattended webview. For mail/calendar commands, the Outlook token is
   also checked. The CLI does not retry commands that fail mid-execution.

5. **Token usage by service**:
   - CSA: `Authorization: Bearer {chatsvcagg_token}` (OAuth token)
   - Messages: `Authentication: skypetoken={authz_skype_token}` (exchanged token)
   - MiddleTier: `Authorization: Bearer {oauth_skype_token}` (OAuth token)
   - Outlook REST API: `Authorization: Bearer {outlook_token}` (OAuth token)

## Security Notes

- Profile names are validated (alphanumeric, dash, underscore only).
- URL parameters are percent-encoded.
- Authz URLs are validated against Microsoft/Skype domains.
- API error bodies are truncated to 500 chars.
- JWT tokens are redacted in debug output.
- Token files use 0600 permissions, token directories use 0700.

## Key Constants

- Teams App ID: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`
- Redirect URI: `https://teams.microsoft.com/go`
- Authz endpoint: `https://teams.microsoft.com/api/authsvc/v1.0/authz`
- Outlook resource: `https://outlook.office365.com`
- Outlook API base: `https://outlook.office365.com/api/v2.0/me`

## API Quirks

- The Messages API uses `Authentication:` header (not `Authorization:`).
- The `messagetype` for plain text is `"Text"`, for HTML is `"RichText/Html"`.
- The `clientmessageid` is a Unix timestamp in milliseconds (as string).
- Region is auto-discovered via authz for all commands (including tenants).
  The `--region` CLI flag is only used as a fallback hint if authz hasn't
  been called or fails.
- The CSA API returns large JSON (~4MB for 700+ chats). Models use
  `serde_json::Value` for flexibility since field shapes vary.
- The MT `/users/me` endpoint doesn't exist. `get_me()` extracts the
  UPN from the JWT and calls `get_user(email)`.

## Conversation IDs

- Channel threads: `19:{uuid}@thread.tacv2`
- Chat threads: `19:{uuid}@thread.v2`
- Self-chat: `19:{uuid1}_{uuid2}@unq.gbl.spaces` (often hidden)
- User MRIs: `8:orgid:{azure-ad-object-id}`

## Tracing

`RUST_LOG` env var controls log level. All logs go to stderr; stdout is
reserved for command output.

## Adding New Commands

1. Add clap structs in `src/cli/<command>.rs`
2. Add the variant to `Commands` enum in `src/cli/mod.rs`
3. Add the handler dispatch in `src/main.rs` `run()` match
4. API calls go in `src/api/`. Use `HttpClient::execute_with_retry()`.
5. Run `cargo fmt && cargo clippy --all-targets -- -D warnings`
