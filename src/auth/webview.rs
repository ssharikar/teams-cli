use std::sync::{Arc, Mutex};

use tao::event::{Event, WindowEvent};
use tao::event_loop::{ControlFlow, EventLoopBuilder};
use tao::platform::run_return::EventLoopExtRunReturn;
use tao::window::WindowBuilder;
use wry::WebViewBuilder;

#[cfg(target_os = "macos")]
mod macos_menu {
    use cocoa::appkit::{NSApp, NSApplication, NSMenu, NSMenuItem};
    use cocoa::base::{nil, selector};
    use cocoa::foundation::{NSAutoreleasePool, NSString};
    use objc::runtime::Sel;

    fn menu_item(title: &str, action: Sel, key: &str) -> cocoa::base::id {
        unsafe {
            NSMenuItem::alloc(nil).initWithTitle_action_keyEquivalent_(
                NSString::alloc(nil).init_str(title),
                action,
                NSString::alloc(nil).init_str(key),
            )
        }
    }

    /// Create a native Edit menu with Undo, Redo, Cut, Copy, Paste, Select All.
    /// This enables Cmd+C/V/X/A keyboard shortcuts in the webview.
    pub fn install_edit_menu() {
        unsafe {
            let _pool = NSAutoreleasePool::new(nil);
            let app = NSApp();

            let menu_bar = NSMenu::new(nil).autorelease();

            // App menu (required placeholder)
            let app_menu_item = NSMenuItem::new(nil).autorelease();
            let app_menu = NSMenu::new(nil).autorelease();
            app_menu_item.setSubmenu_(app_menu);
            menu_bar.addItem_(app_menu_item);

            // Edit menu
            let edit_menu_item = NSMenuItem::new(nil).autorelease();
            let edit_menu = NSMenu::alloc(nil)
                .initWithTitle_(NSString::alloc(nil).init_str("Edit"))
                .autorelease();

            edit_menu.addItem_(menu_item("Undo", selector("undo:"), "z"));
            edit_menu.addItem_(menu_item("Redo", selector("redo:"), "Z"));
            edit_menu.addItem_(NSMenuItem::separatorItem(nil));
            edit_menu.addItem_(menu_item("Cut", selector("cut:"), "x"));
            edit_menu.addItem_(menu_item("Copy", selector("copy:"), "c"));
            edit_menu.addItem_(menu_item("Paste", selector("paste:"), "v"));
            edit_menu.addItem_(menu_item("Select All", selector("selectAll:"), "a"));

            edit_menu_item.setSubmenu_(edit_menu);
            menu_bar.addItem_(edit_menu_item);

            app.setMainMenu_(menu_bar);
        }
    }
}

use super::token::{
    extract_tenant_id, TokenInfo, TokenSet, TokenType, CHATSVCAGG_RESOURCE, MICROSOFT_TENANT_ID,
    OUTLOOK_RESOURCE, REDIRECT_URI, SKYPE_RESOURCE, TEAMS_APP_ID,
};
use crate::error::TeamsError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AuthPhase {
    Teams,
    Skype,
    ChatSvcAgg,
    Outlook,
    Done,
}

struct AuthState {
    phase: AuthPhase,
    teams_token: Option<String>,
    skype_token: Option<String>,
    chatsvcagg_token: Option<String>,
    outlook_token: Option<String>,
    tenant_id: String,
    redirect_count: u32,
    expected_state: Option<String>,
}

fn build_auth_url(
    tenant: &str,
    response_type: &str,
    resource: Option<&str>,
    client_id: &str,
    prompt: Option<&str>,
) -> (String, String) {
    let state = uuid::Uuid::new_v4().to_string();
    let nonce = uuid::Uuid::new_v4().to_string();
    let client_request_id = uuid::Uuid::new_v4().to_string();

    let mut url = format!(
        "https://login.microsoftonline.com/{tenant}/oauth2/authorize\
         ?response_type={response_type}\
         &client_id={client_id}\
         &redirect_uri={REDIRECT_URI}\
         &state={state}\
         &client-request-id={client_request_id}\
         &nonce={nonce}\
         &x-client-SKU=Js\
         &x-client-Ver=1.0.9"
    );

    if let Some(res) = resource {
        url.push_str(&format!("&resource={res}"));
    }

    // Only set on the very first navigation of a login flow. Without it,
    // Microsoft silently reauthenticates as whatever account still has an
    // active SSO session (e.g. from a previous `auth login` under a
    // different --profile), instead of letting the person pick — which is
    // how `auth login --profile cms` ended up signed in as the Samtek
    // account. The 3 follow-up token requests in the same flow intentionally
    // omit this so they stay silent and reuse the account just chosen here.
    if let Some(p) = prompt {
        url.push_str(&format!("&prompt={p}"));
    }

    (url, state)
}

fn extract_token_from_fragment(url_str: &str) -> Option<(String, bool, Option<String>)> {
    let fragment = url_str.split('#').nth(1)?;
    let params = url::form_urlencoded::parse(fragment.as_bytes());
    let mut token = None;
    let mut is_id_token = false;
    let mut state_param = None;

    for (key, value) in params {
        match key.as_ref() {
            "id_token" => {
                token = Some(value.to_string());
                is_id_token = true;
            }
            "access_token" => {
                if token.is_none() {
                    token = Some(value.to_string());
                }
            }
            "state" => {
                state_param = Some(value.to_string());
            }
            _ => {}
        }
    }

    token.map(|t| (t, is_id_token, state_param))
}

fn is_allowed_domain(url: &str) -> bool {
    let allowed = [
        "login.microsoftonline.com",
        "login.microsoft.com",
        "login.live.com",
        "teams.microsoft.com",
        "aadcdn.msauth.net",
        "aadcdn.msftauth.net",
        // Federated identity providers (SSO redirects during OAuth)
        "sso.godaddy.com",
        "sso.secureserver.net",
    ];
    if let Ok(parsed) = url::Url::parse(url) {
        if let Some(host) = parsed.host_str() {
            return allowed
                .iter()
                .any(|d| host == *d || host.ends_with(&format!(".{d}")));
        }
    }
    false
}

/// Run the webview-based 4-token OAuth2 flow (Teams, Skype, ChatSvcAgg, Outlook).
/// All tokens use implicit grant via the Teams app ID.
/// Must be called from the main thread. Returns the token set on success.
pub fn webview_login(initial_tenant: &str, profile: &str) -> crate::error::Result<TokenSet> {
    let (initial_url, initial_state) = build_auth_url(
        initial_tenant,
        "id_token",
        None,
        TEAMS_APP_ID,
        Some("select_account"),
    );

    let state = Arc::new(Mutex::new(AuthState {
        phase: AuthPhase::Teams,
        teams_token: None,
        skype_token: None,
        chatsvcagg_token: None,
        outlook_token: None,
        tenant_id: initial_tenant.to_string(),
        redirect_count: 0,
        expected_state: Some(initial_state),
    }));

    let profile = profile.to_string();
    let result: Arc<Mutex<Option<crate::error::Result<TokenSet>>>> = Arc::new(Mutex::new(None));

    // Install native Edit menu on macOS so Cmd+C/V/X/A work in the webview
    #[cfg(target_os = "macos")]
    macos_menu::install_edit_menu();

    let mut event_loop = EventLoopBuilder::<String>::with_user_event().build();
    let proxy = event_loop.create_proxy();

    let window = WindowBuilder::new()
        .with_title("Teams CLI - Sign In")
        .with_inner_size(tao::dpi::LogicalSize::new(500.0, 700.0))
        .with_focused(true)
        .build(&event_loop)
        .expect("failed to create window");

    let state_for_nav = state.clone();
    let proxy_for_nav = proxy.clone();

    let webview = WebViewBuilder::new()
        .with_url(&initial_url)
        // Belt-and-suspenders alongside prompt=select_account above: an
        // ephemeral, non-persistent session for this login window means no
        // SSO cookie from a previous `auth login` (possibly a different
        // --profile/account) can be reused, by construction.
        .with_incognito(true)
        .with_navigation_handler(move |url: String| {
            if !url.starts_with("https://teams.microsoft.com/go") {
                if !is_allowed_domain(&url) {
                    tracing::warn!("blocked navigation to disallowed domain: {url}");
                    return false;
                }
                return true;
            }

            let mut auth = state_for_nav.lock().unwrap_or_else(|e| e.into_inner());
            auth.redirect_count += 1;

            if auth.redirect_count > 12 {
                let _ = proxy_for_nav.send_event("error:too many redirects".into());
                return false;
            }

            let Some((token, is_id_token, returned_state)) = extract_token_from_fragment(&url)
            else {
                return false;
            };

            if let Some(ref expected) = auth.expected_state {
                match returned_state {
                    Some(ref rs) if rs != expected => {
                        tracing::warn!("OAuth state mismatch: expected={expected}, got={rs}");
                        return false;
                    }
                    None => {
                        tracing::warn!("OAuth response missing state parameter");
                        return false;
                    }
                    _ => {}
                }
            }

            match auth.phase {
                AuthPhase::Teams => {
                    if is_id_token {
                        if let Ok(Some(tid)) = extract_tenant_id(&token) {
                            if tid == MICROSOFT_TENANT_ID {
                                auth.tenant_id = "common".to_string();
                            } else {
                                auth.tenant_id = tid;
                            }
                        }
                        auth.teams_token = Some(token);
                        auth.phase = AuthPhase::Skype;
                        let (next_url, next_state) = build_auth_url(
                            &auth.tenant_id,
                            "token",
                            Some(SKYPE_RESOURCE),
                            TEAMS_APP_ID,
                            None,
                        );
                        auth.expected_state = Some(next_state);
                        let _ = proxy_for_nav.send_event(format!("navigate:{next_url}"));
                    }
                }
                AuthPhase::Skype => {
                    auth.skype_token = Some(token);
                    auth.phase = AuthPhase::ChatSvcAgg;
                    let (next_url, next_state) = build_auth_url(
                        &auth.tenant_id,
                        "token",
                        Some(CHATSVCAGG_RESOURCE),
                        TEAMS_APP_ID,
                        None,
                    );
                    auth.expected_state = Some(next_state);
                    let _ = proxy_for_nav.send_event(format!("navigate:{next_url}"));
                }
                AuthPhase::ChatSvcAgg => {
                    auth.chatsvcagg_token = Some(token);
                    auth.phase = AuthPhase::Outlook;
                    let (next_url, next_state) = build_auth_url(
                        &auth.tenant_id,
                        "token",
                        Some(OUTLOOK_RESOURCE),
                        TEAMS_APP_ID,
                        None,
                    );
                    auth.expected_state = Some(next_state);
                    let _ = proxy_for_nav.send_event(format!("navigate:{next_url}"));
                }
                AuthPhase::Outlook => {
                    auth.outlook_token = Some(token);
                    auth.phase = AuthPhase::Done;
                    auth.expected_state = None;
                    let _ = proxy_for_nav.send_event("done".into());
                }
                AuthPhase::Done => {}
            }

            false // Always block redirect to teams.microsoft.com/go
        })
        .build(&window)
        .expect("failed to create webview");

    let state_for_loop = state.clone();
    let result_for_loop = result.clone();

    event_loop.run_return(move |event, _, control_flow| {
        *control_flow = ControlFlow::Wait;

        match event {
            Event::UserEvent(ref msg) => {
                if let Some(url) = msg.strip_prefix("navigate:") {
                    if let Err(e) = webview.load_url(url) {
                        *result_for_loop.lock().unwrap_or_else(|e| e.into_inner()) = Some(Err(
                            TeamsError::AuthError(format!("failed to navigate: {e}")),
                        ));
                        *control_flow = ControlFlow::Exit;
                    }
                } else if msg == "done" {
                    let auth = state_for_loop.lock().unwrap_or_else(|e| e.into_inner());
                    let outcome = build_token_set(&auth, &profile).and_then(|token_set| {
                        crate::auth::keyring::store_tokens(&profile, &token_set)?;
                        Ok(token_set)
                    });
                    *result_for_loop.lock().unwrap_or_else(|e| e.into_inner()) = Some(outcome);
                    *control_flow = ControlFlow::Exit;
                } else if let Some(err_msg) = msg.strip_prefix("error:") {
                    *result_for_loop.lock().unwrap_or_else(|e| e.into_inner()) =
                        Some(Err(TeamsError::AuthError(err_msg.to_string())));
                    *control_flow = ControlFlow::Exit;
                }
            }
            Event::WindowEvent {
                event: WindowEvent::CloseRequested,
                ..
            } => {
                let auth = state_for_loop.lock().unwrap_or_else(|e| e.into_inner());
                if auth.phase == AuthPhase::Done {
                    let outcome = build_token_set(&auth, &profile).and_then(|token_set| {
                        crate::auth::keyring::store_tokens(&profile, &token_set)?;
                        Ok(token_set)
                    });
                    *result_for_loop.lock().unwrap_or_else(|e| e.into_inner()) = Some(outcome);
                } else {
                    *result_for_loop.lock().unwrap_or_else(|e| e.into_inner()) =
                        Some(Err(TeamsError::AuthError("login cancelled by user".into())));
                }
                *control_flow = ControlFlow::Exit;
            }
            _ => {}
        }
    });

    // Extract the result from the event loop
    let outcome = result
        .lock()
        .unwrap_or_else(|e| e.into_inner())
        .take()
        .unwrap_or_else(|| Err(TeamsError::AuthError("login did not complete".into())));

    outcome
}

fn build_token_set(auth: &AuthState, profile: &str) -> crate::error::Result<TokenSet> {
    let teams_raw = auth
        .teams_token
        .as_ref()
        .ok_or_else(|| TeamsError::AuthError("missing teams token".into()))?;
    let skype_raw = auth
        .skype_token
        .as_ref()
        .ok_or_else(|| TeamsError::AuthError("missing skype token".into()))?;
    let chatsvcagg_raw = auth
        .chatsvcagg_token
        .as_ref()
        .ok_or_else(|| TeamsError::AuthError("missing chatsvcagg token".into()))?;

    let outlook = auth
        .outlook_token
        .as_ref()
        .map(|raw| TokenInfo::from_jwt(raw, TokenType::AccessToken))
        .transpose()?;

    Ok(TokenSet {
        teams: TokenInfo::from_jwt(teams_raw, TokenType::IdToken)?,
        skype: TokenInfo::from_jwt(skype_raw, TokenType::AccessToken)?,
        chatsvcagg: TokenInfo::from_jwt(chatsvcagg_raw, TokenType::AccessToken)?,
        outlook,
        profile: profile.to_string(),
        tenant_id: auth.tenant_id.clone(),
    })
}
