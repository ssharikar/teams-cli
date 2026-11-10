# teams-cli

CLI for Microsoft Teams and Outlook using internal Microsoft APIs (not the Microsoft Graph API).

## Prerequisites

- **macOS**: No additional dependencies (WebKit via WKWebView).
- **Linux**: `libwebkit2gtk-4.1-dev`, `libgtk-3-dev`.
- **Windows**: WebView2 (pre-installed on Windows 10+).

## Install

```sh
cargo install --path .
```

## Authentication

Teams CLI acquires four OAuth2 tokens via a native webview that opens the
Microsoft login page (Teams, Skype, ChatSvcAgg, Outlook). After
authentication, a token exchange via the authz endpoint provides a
messaging-capable Skype token and auto-discovers the correct regional API
endpoints.

```sh
# Login via native webview
teams auth login

# Check token status
teams auth status

# List every profile with stored credentials (e.g. multiple accounts/tenants)
teams auth list

# Export a token for scripting
teams auth token skype
teams auth token outlook

# Clear credentials
teams auth logout
```

Tokens are stored as JSON files at `<config-dir>/teams-cli/tokens/<profile>.json`
(file mode `0600`, directory mode `0700`). The config directory is
platform-dependent: `~/Library/Application Support` on macOS,
`~/.config` on Linux. Run `teams config path` to see the resolved location.

### Auto-login

If tokens are missing or expired, the CLI will attempt to authenticate before
running the command (unless `--no-auto-login` is set). The CLI does not retry
commands that fail mid-execution.

## Usage

### Teams

```sh
teams team list
teams team get <team-id>
```

### Channels

```sh
teams channel list <team-id>
teams channel get <team-id> <channel-id>
teams channel pinned
```

### Chats

```sh
teams chat list
teams chat list --all   # include hidden chats
teams chat get <chat-id>
```

### Messages

```sh
teams message list <conversation-id> --limit 50
teams message send <conversation-id> --body "Hello from the CLI"
echo "piped message" | teams message send <conversation-id> --stdin
teams message get <conversation-id> <message-id>

# Edit or delete a message you previously sent
teams message edit <conversation-id> <message-id> --body "corrected text"
teams message delete <conversation-id> <message-id>
```

### Users

```sh
teams user me
teams user get user@example.com
teams user search "8:orgid:mri-1,8:orgid:mri-2"
```

### Email (Outlook)

```sh
# List recent inbox messages
teams mail list
teams mail list --since 1h
teams mail list --since 24h --unread
teams mail list --folder "Sent Items" --limit 10

# Read a specific email
teams mail read <message-id>

# Send an email
teams mail send --to user@example.com --subject "Hello" --body "Message body"
teams mail send --to user@example.com --cc other@example.com --subject "FYI" --body "See attached"
teams mail send --to user@example.com --bcc user1@example.com,user2@example.com --subject "FYI" --body "See attached"
echo "piped body" | teams mail send --to user@example.com --subject "Test" --stdin
teams mail send --to user@example.com --subject "HTML" --body "<h1>Rich</h1>" --html

# Save an email to Drafts without sending it; --to/--cc/--bcc take a repeatable
# flag or a comma-separated list (handles large Bcc lists, e.g. ~180 recipients)
teams mail draft --to user@example.com --bcc a@example.com,b@example.com --subject "Draft" --body "Review before sending"

# Send a draft created above, once it's ready
teams mail send-draft <message-id>

# Permanently delete a message (e.g. an unwanted draft)
teams mail delete <message-id>

# Search emails
teams mail search "quarterly report" --limit 10
```

### Calendar (Outlook)

```sh
# List upcoming events (default: next 7 days)
teams calendar list
teams calendar list --from today --to +3d
teams calendar list --from 2026-05-01 --to 2026-05-31

# Get event details
teams calendar get <event-id>

# Create a meeting
teams calendar create --subject "Standup" --start 2026-04-28T09:00:00 --end 2026-04-28T09:30:00
teams calendar create --subject "Review" --start 2026-04-28T14:00:00 --end 2026-04-28T15:00:00 \
  --attendees alice@example.com --attendees bob@example.com --online --location "Room 1"
```

Outlook commands use lazy token acquisition. The first time you run a mail or
calendar command, you'll be prompted to authenticate via device code flow for
the Outlook API. The token is cached alongside your Teams tokens.

### Presence

```sh
# Get your current presence status
teams presence get

# Set presence to online / offline
teams presence online
teams presence offline

# Other statuses
teams presence away
teams presence busy
teams presence dnd

# Set a custom availability
teams presence set BeRightBack
```

Valid availability values: `Available`, `Busy`, `DoNotDisturb`, `BeRightBack`, `Away`, `Offline`.

#### Scheduling with cron

Automate your presence with crontab (`crontab -e`):

```cron
# Go online at 9 AM Mon-Fri
0 9 * * 1-5 /path/to/teams presence online --no-auto-login

# Go offline at 5 PM Mon-Fri
0 17 * * 1-5 /path/to/teams presence offline --no-auto-login
```

### Tenants

```sh
teams tenant list
teams tenant domains
```

### Configuration

```sh
teams config init
teams config show
teams config set default.region amer
teams config path
```

### Shell Completions

```sh
teams completions bash > ~/.bash_completion.d/teams
teams completions zsh > ~/.zfunc/_teams
teams completions fish > ~/.config/fish/completions/teams.fish
```

## Output Formats

```sh
# Auto-detect (human for TTY, JSON for pipes)
teams team list

# Force JSON envelope
teams team list --output json

# Plain text (tab-separated, for scripting)
teams team list --output plain
```

Unrecognized format strings are rejected with an error.

## Global Options

| Flag | Description |
|------|-------------|
| `--output <format>` | json, human, plain (auto-detect by default) |
| `--region <region>` | API region hint: emea, amer, apac. Region is auto-detected via authz for all commands; this flag is only used as a fallback if authz fails. |
| `--profile <name>` | Named credential profile (alphanumeric, dash, underscore only) |
| `--timeout <secs>` | Request timeout (default: 30) |
| `--retry <count>` | Max retry attempts (default: 3) |
| `--no-auto-login` | Skip automatic authentication when tokens are missing/expired |
| `--no-color` | Disable ANSI color output (also respects `NO_COLOR` env var) |
| `-v` / `-vv` / `-vvv` | Verbosity: info / debug / trace |
| `-q` / `--quiet` | Suppress non-essential output |

## Environment Variables

| Variable | Description |
|----------|-------------|
| `TEAMS_CLI_TEAMS_TOKEN` | Override Teams JWT |
| `TEAMS_CLI_SKYPE_TOKEN` | Override Skype JWT |
| `TEAMS_CLI_CHATSVCAGG_TOKEN` | Override ChatSvcAgg JWT |
| `TEAMS_CLI_OUTLOOK_TOKEN` | Override Outlook JWT (optional, for mail/calendar) |
| `NO_COLOR` | Disable ANSI color output (any value) |
| `RUST_LOG` | Tracing filter (e.g. `debug`) |

All three token env vars (`TEAMS_CLI_TEAMS_TOKEN`, `TEAMS_CLI_SKYPE_TOKEN`,
`TEAMS_CLI_CHATSVCAGG_TOKEN`) must be set together to override file-based auth.
Setting only one or two is not sufficient.

## Configuration File

Config lives at `<config-dir>/teams-cli/config.toml`. Initialize with `teams config init`.
Run `teams config path` to see the resolved location.

```toml
[default]
profile = "default"     # Default profile name
region = "emea"         # Default region: emea, amer, apac

[output]
format = "auto"         # Output format: auto, json, human, plain
color = true            # Enable ANSI colors

[network]
timeout = 30            # Request timeout in seconds
max_retries = 3         # Max retry attempts
retry_backoff_base = 2  # Exponential backoff base in seconds

[profiles.myorg]
tenant_id = "common"    # Azure AD tenant ID or "common"
```

Config defaults (`default.profile`, `default.region`) are used when CLI flags
are not explicitly provided. The clap defaults ("default" for profile, "emea"
for region) act as sentinels -- if you haven't overridden them on the command
line, the config file values are used instead.

### Config set keys

| Key | Values | Description |
|-----|--------|-------------|
| `default.profile` | string | Default profile name |
| `default.region` | `emea`, `amer`, `apac` | Default API region |
| `output.format` | `auto`, `json`, `human`, `plain` | Default output format |
| `output.color` | `true`, `false` | Enable ANSI colors |
| `network.timeout` | integer (seconds) | Request timeout |
| `network.max_retries` | integer | Max retry attempts |

## Exit Codes

| Code | Meaning | Error Codes |
|------|---------|-------------|
| 0 | Success | - |
| 1 | General error | API_ERROR, UNKNOWN |
| 2 | Invalid input | INVALID_INPUT |
| 3 | Auth failure | AUTH_FAILED, AUTH_TOKEN_EXPIRED |
| 4 | Permission denied | PERMISSION_DENIED |
| 5 | Not found | NOT_FOUND |
| 6 | Rate limited | RATE_LIMITED |
| 7 | Network error | NETWORK_ERROR |
| 8 | Server error (5xx) | SERVER_ERROR |
| 10 | Config/keyring error | CONFIG_ERROR, KEYRING_ERROR |

HTTP status codes in API errors map to specific exit codes: 401 -> 3, 403 -> 4,
404 -> 5, 429 -> 6, 5xx -> 8, other -> 1. JSON output includes `error.code`
with the symbolic error code.

## API Services

The CLI communicates with three Microsoft Teams backend services, discovered
dynamically via the authz token exchange:

- **Chat Service Aggregator (CSA)** -- teams, channels, chats listing
- **Messages Service** -- message read/write (uses the authz-exchanged Skype token)
- **MiddleTier (MT)** -- user profiles, tenants, domains
- **Presence Service** -- presence/availability status (at `presence.teams.microsoft.com`)
- **Outlook REST API v2.0** -- email and calendar operations (at `outlook.office365.com/api/v2.0`)

Teams commands use the same internal APIs as the official Teams client.
Outlook commands use the Outlook REST API v2.0 with the same app identity,
avoiding the need for Microsoft Graph API registration.

## License

MIT -- see [LICENSE](LICENSE).
