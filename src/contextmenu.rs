//! MDView's right-click menu (#59): a standard Windows popup menu with the
//! keys on the right, instead of WebView2's own menu.

use windows::Win32::Foundation::{HWND, POINT};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, DestroyMenu, MF_GRAYED, MF_SEPARATOR, MF_STRING, SetForegroundWindow,
    TPM_RETURNCMD, TPM_RIGHTBUTTON, TrackPopupMenu,
};
use windows::core::PCWSTR;

use crate::shortcuts::{MenuCommand, MenuItem};
use webview2_com::Microsoft::Web::WebView2::Win32::{ICoreWebView2, ICoreWebView2_11};
use windows::Win32::Graphics::Gdi::ClientToScreen;
use windows::core::Interface;

/// A right-click (or the menu key) in the page.
#[derive(Debug, Clone)]
pub struct Request {
    /// The link's URL when the click was on a link.
    pub link: Option<String>,
    pub has_selection: bool,
    /// Where to show the menu, in screen coordinates.
    pub screen: POINT,
}

/// Replace WebView2's menu: every right-click is reported to `on_request`,
/// which shows MDView's menu later (never inside this WebView2 callback).
/// While a panel is open (`panel_open`) there is no menu at all. `host` is
/// the window the WebView fills.
pub fn register(
    webview: &ICoreWebView2,
    host: HWND,
    panel_open: impl Fn() -> bool + 'static,
    on_request: impl Fn(Request) + 'static,
) -> Result<(), String> {
    let webview11: ICoreWebView2_11 = webview
        .cast()
        .map_err(|_| "this version of the WebView2 Runtime has no context menu events".to_string())?;
    let handler = webview2_com::ContextMenuRequestedEventHandler::create(Box::new(move |_sender, args| {
        let Some(args) = args else { return Ok(()) };
        unsafe {
            args.SetHandled(true)?;
            if panel_open() {
                return Ok(());
            }
            let mut link = None;
            let mut has_selection = false;
            if let Ok(target) = args.ContextMenuTarget() {
                let mut is_link = windows::core::BOOL(0);
                if target.HasLinkUri(&mut is_link).is_ok() && is_link.as_bool() {
                    let mut uri = windows::core::PWSTR::null();
                    if target.LinkUri(&mut uri).is_ok() && !uri.is_null() {
                        link = uri.to_string().ok().filter(|u| !u.is_empty());
                        windows::Win32::System::Com::CoTaskMemFree(Some(uri.0 as *const _));
                    }
                }
                let mut selection = windows::core::BOOL(0);
                has_selection = target.HasSelection(&mut selection).is_ok() && selection.as_bool();
            }
            let mut screen = POINT::default();
            let _ = args.Location(&mut screen);
            let _ = ClientToScreen(host, &mut screen);
            on_request(Request { link, has_selection, screen });
        }
        Ok(())
    }));
    let mut token = 0i64;
    unsafe { webview11.add_ContextMenuRequested(&handler, &mut token) }.map_err(|e| e.message().to_string())
}

/// Show the menu at `screen` (screen coordinates) and return the chosen
/// command; `None` when the user closes it without choosing.
pub fn show(owner: HWND, screen: POINT, items: &[Option<MenuItem>]) -> Option<MenuCommand> {
    unsafe {
        let menu = CreatePopupMenu().ok()?;
        for item in items {
            match item {
                None => {
                    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());
                }
                Some(item) => {
                    let text: Vec<u16> = item.label.encode_utf16().chain(std::iter::once(0)).collect();
                    let flags = if item.enabled { MF_STRING } else { MF_STRING | MF_GRAYED };
                    let _ = AppendMenuW(menu, flags, item.command as usize, PCWSTR(text.as_ptr()));
                }
            }
        }
        // Needed so the menu closes when the user clicks elsewhere.
        let _ = SetForegroundWindow(owner);
        let chosen = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen.x, screen.y, None, owner, None);
        let _ = DestroyMenu(menu);
        MenuCommand::from_id(chosen.0 as u32)
    }
}
