//! Where a link goes, shown in a standard Windows tooltip while the mouse is
//! over it (#60). Shared by the standalone viewer and the TC plugin.
//!
//! The text comes from MDView, from the same decision a click makes
//! (`links::classify_url`), never from the document: a link's own `title`
//! can say anything, and a document's CSS can hide or cover anything in the
//! page, but not a native tooltip. The page script only reports which link
//! is under the mouse.

use std::cell::RefCell;
use std::collections::HashMap;

use windows::Win32::Foundation::{HWND, LPARAM, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetMonitorInfoW, MONITOR_DEFAULTTONEAREST, MONITORINFO, MonitorFromPoint};
use windows::Win32::UI::Controls::{
    SetWindowTheme, TOOLTIPS_CLASSW, TTF_ABSOLUTE, TTF_TRACK, TTF_TRANSPARENT, TTM_ADDTOOLW,
    TTM_SETMAXTIPWIDTH, TTM_TRACKACTIVATE, TTM_TRACKPOSITION, TTM_UPDATETIPTEXTW, TTS_ALWAYSTIP,
    TTS_NOPREFIX, TTTOOLINFOW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetCursorPos, GetWindowRect, KillTimer, SPI_GETMOUSEHOVERTIME, SendMessageW,
    SetTimer, SystemParametersInfoW, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS, WINDOW_STYLE,
    WS_EX_TOPMOST, WS_POPUP,
};
use windows::core::{PCWSTR, PWSTR};

use crate::links::LinkAction;

/// The text for a link: what a click on it would open. `href` is the link's
/// attribute as written, `action` the click's decision for its full URL.
pub fn describe(href: &str, action: &LinkAction) -> String {
    if href.starts_with('#') {
        return href.to_string();
    }
    match action {
        LinkAction::OpenExternal(url) => url.clone(),
        LinkAction::OpenMarkdown(path) | LinkAction::ConfirmOpenFile(path) => path.display().to_string(),
        LinkAction::ConfirmOpenNetwork { path, .. } => path.display().to_string(),
        LinkAction::Program { path, .. } => format!("{} (program: Ctrl+click to run)", path.display()),
        LinkAction::Refuse { target, .. } => format!("{target} (not opened)"),
    }
}

/// `#heading` part of a URL, if any, so a link to a heading in another
/// document shows `C:\docs\b.md#heading`.
pub fn fragment_of(url: &str) -> Option<&str> {
    url.find('#').map(|i| &url[i..]).filter(|f| f.len() > 1)
}

/// The text for a link, with the heading it jumps to.
pub fn describe_url(url: &str, href: &str, action: &LinkAction) -> String {
    let text = describe(href, action);
    match (action, fragment_of(url)) {
        (LinkAction::OpenMarkdown(_), Some(fragment)) if !href.starts_with('#') => format!("{text}{fragment}"),
        _ => text,
    }
}

thread_local! {
    /// One tooltip per host window (the plugin can have several Lister windows),
    /// keyed by the host window.
    static TIPS: RefCell<HashMap<isize, HWND>> = RefCell::new(HashMap::new());
    /// Where to show the tooltip instead of next to the mouse (`show_at`).
    static ANCHORS: RefCell<HashMap<isize, POINT>> = RefCell::new(HashMap::new());
}

const TOOL_ID: usize = 1;
const SHOW_TIMER: usize = 0x4D44; // "MD"

fn tool_info(owner: HWND, text: *mut u16) -> TTTOOLINFOW {
    TTTOOLINFOW {
        // Without lpReserved: accepted by every comctl32 version, also when
        // the host process has no common-controls 6 manifest.
        cbSize: std::mem::offset_of!(TTTOOLINFOW, lpReserved) as u32,
        // Transparent: the mouse goes through to the page, so showing the
        // tooltip never counts as leaving the link.
        uFlags: TTF_TRACK | TTF_ABSOLUTE | TTF_TRANSPARENT,
        hwnd: owner,
        uId: TOOL_ID,
        lpszText: PWSTR(text),
        ..Default::default()
    }
}

fn create(owner: HWND, dark: bool) -> Option<HWND> {
    let window = unsafe {
        CreateWindowExW(
            WS_EX_TOPMOST,
            TOOLTIPS_CLASSW,
            PCWSTR::null(),
            WS_POPUP | WINDOW_STYLE(TTS_NOPREFIX | TTS_ALWAYSTIP),
            0,
            0,
            0,
            0,
            Some(owner),
            None,
            None,
            None,
        )
    }
    .ok()?;
    if dark {
        let _ = unsafe { SetWindowTheme(window, windows::core::w!("DarkMode_Explorer"), PCWSTR::null()) };
    }
    let mut empty = [0u16; 1];
    let info = tool_info(owner, empty.as_mut_ptr());
    unsafe {
        SendMessageW(window, TTM_ADDTOOLW, None, Some(LPARAM(&info as *const _ as isize)));
    }
    Some(window)
}

/// Long targets without spaces (paths, URLs) would make one line wider than
/// the screen: break after a `\` or `/` once a line is this long, or anywhere
/// in a longer piece without one. Display only; the text is not changed.
const LINE_CHARS: usize = 100;

pub fn wrap_for_display(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + text.len() / LINE_CHARS * 2);
    let mut line = 0;
    for c in text.chars() {
        out.push(c);
        line += 1;
        if (line >= LINE_CHARS && matches!(c, '\\' | '/')) || line >= LINE_CHARS * 3 / 2 {
            out.push_str("\r\n");
            line = 0;
        }
    }
    out
}

/// Show `text` for the link under the mouse, after the system's hover delay.
/// Replaces what was shown before.
pub fn show(owner: HWND, text: &str, dark: bool) {
    ANCHORS.with(|a| a.borrow_mut().remove(&(owner.0 as isize)));
    show_tip(owner, text, dark);
}

/// Show `text` at `at` (screen coordinates), for example next to a menu
/// item (#44), after the hover delay.
#[allow(dead_code)] // standalone viewer only
pub fn show_at(owner: HWND, text: &str, dark: bool, at: POINT) {
    ANCHORS.with(|a| a.borrow_mut().insert(owner.0 as isize, at));
    show_tip(owner, text, dark);
}

fn show_tip(owner: HWND, text: &str, dark: bool) {
    let key = owner.0 as isize;
    let existing = TIPS.with(|t| t.borrow().get(&key).copied());
    let Some(window) = existing.or_else(|| create(owner, dark)) else { return };
    TIPS.with(|t| t.borrow_mut().insert(key, window));
    hide_window(window);
    let mut wide: Vec<u16> = wrap_for_display(text).encode_utf16().chain(std::iter::once(0)).collect();
    let info = tool_info(owner, wide.as_mut_ptr());
    unsafe {
        SendMessageW(window, TTM_UPDATETIPTEXTW, None, Some(LPARAM(&info as *const _ as isize)));
    }
    let mut delay: u32 = 400;
    let _ = unsafe {
        SystemParametersInfoW(
            SPI_GETMOUSEHOVERTIME,
            0,
            Some(&mut delay as *mut u32 as *mut _),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
    };
    unsafe { SetTimer(Some(window), SHOW_TIMER, delay, Some(show_now)) };
}

fn owner_of(window: HWND) -> Option<HWND> {
    TIPS.with(|t| t.borrow().iter().find(|(_, tip)| **tip == window).map(|(k, _)| HWND(*k as *mut _)))
}

fn track_position(window: HWND, x: i32, y: i32) {
    let position = ((y as u32 & 0xFFFF) << 16) | (x as u32 & 0xFFFF);
    unsafe { SendMessageW(window, TTM_TRACKPOSITION, None, Some(LPARAM(position as isize))) };
}

/// The hover delay is over: show the tooltip below-right of the mouse, then
/// move it back onto the screen if it sticks out. (Its size is read from the
/// shown window: TTM_GETBUBBLESIZE is not reliable before the first show and
/// fails in some host processes.)
unsafe extern "system" fn show_now(window: HWND, _msg: u32, id: usize, _time: u32) {
    unsafe {
        let _ = KillTimer(Some(window), id);
        let Some(owner) = owner_of(window) else { return };
        let anchor = ANCHORS.with(|a| a.borrow().get(&(owner.0 as isize)).copied());
        let mut cursor = POINT::default();
        if let Some(at) = anchor {
            cursor = at;
        } else if GetCursorPos(&mut cursor).is_err() {
            return;
        }
        let monitor = MonitorFromPoint(cursor, MONITOR_DEFAULTTONEAREST);
        let mut info = MONITORINFO { cbSize: std::mem::size_of::<MONITORINFO>() as u32, ..Default::default() };
        if !GetMonitorInfoW(monitor, &mut info).as_bool() {
            return;
        }
        let work = info.rcWork;
        let max_width = ((work.right - work.left) * 2 / 3).max(200);
        SendMessageW(window, TTM_SETMAXTIPWIDTH, None, Some(LPARAM(max_width as isize)));
        // Next to the mouse; exactly at the anchor when there is one.
        let (dx, dy) = if anchor.is_some() { (0, 0) } else { (12, 20) };
        track_position(window, cursor.x + dx, cursor.y + dy);
        let mut empty = [0u16; 1];
        let tool = tool_info(owner, empty.as_mut_ptr());
        SendMessageW(window, TTM_TRACKACTIVATE, Some(WPARAM(1)), Some(LPARAM(&tool as *const _ as isize)));
        let mut shown = RECT::default();
        if GetWindowRect(window, &mut shown).is_ok() {
            let (width, height) = (shown.right - shown.left, shown.bottom - shown.top);
            let mut x = shown.left;
            let mut y = shown.top;
            if shown.right > work.right {
                x = (work.right - width).max(work.left);
            }
            if shown.bottom > work.bottom {
                y = (cursor.y - height - 4).max(work.top);
            }
            if (x, y) != (shown.left, shown.top) {
                track_position(window, x, y);
            }
        }
    }
}

fn hide_window(window: HWND) {
    unsafe {
        let _ = KillTimer(Some(window), SHOW_TIMER);
        if let Some(owner) = owner_of(window) {
            let mut empty = [0u16; 1];
            let tool = tool_info(owner, empty.as_mut_ptr());
            SendMessageW(window, TTM_TRACKACTIVATE, Some(WPARAM(0)), Some(LPARAM(&tool as *const _ as isize)));
        }
    }
}

/// The mouse left the link, the page scrolled or navigated, or the window
/// lost focus.
pub fn hide(owner: HWND) {
    if let Some(window) = TIPS.with(|t| t.borrow().get(&(owner.0 as isize)).copied()) {
        hide_window(window);
    }
}

/// The host window is going away. In the plugin the tooltip's owner is TC's
/// Lister window (a popup is owned by a top-level window), so it would
/// otherwise outlive the viewer.
pub fn destroy(owner: HWND) {
    if let Some(window) = TIPS.with(|t| t.borrow_mut().remove(&(owner.0 as isize))) {
        unsafe {
            let _ = KillTimer(Some(window), SHOW_TIMER);
            let _ = DestroyWindow(window);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    // #60: the tooltip says what a click would open, not what the link's text
    // or title claims.
    #[test]
    fn test_describe_shows_what_a_click_opens() {
        let web = LinkAction::OpenExternal("https://xn--pypal-4ve.com/login".to_string());
        assert_eq!(describe("https://paypal.com", &web), "https://xn--pypal-4ve.com/login");
        let md = LinkAction::OpenMarkdown(PathBuf::from(r"C:\docs\b.md"));
        assert_eq!(describe("b.md", &md), r"C:\docs\b.md");
        let file = LinkAction::ConfirmOpenFile(PathBuf::from(r"D:\pictures\x.png"));
        assert_eq!(describe("D:/pictures/x.png", &file), r"D:\pictures\x.png");
        let network = LinkAction::ConfirmOpenNetwork {
            path: PathBuf::from(r"\\server\share\x.md"),
            location: r"\\server\share".to_string(),
        };
        assert_eq!(describe("file://server/share/x.md", &network), r"\\server\share\x.md");
        let program = LinkAction::Program { path: PathBuf::from(r"C:\docs\run.exe"), location: None };
        assert_eq!(describe("run.exe", &program), r"C:\docs\run.exe (program: Ctrl+click to run)");
        let refused = LinkAction::Refuse { target: "ms-msdt:/id x".to_string(), reason: crate::links::REFUSED_LINK_TYPE };
        assert_eq!(describe("ms-msdt:/id x", &refused), "ms-msdt:/id x (not opened)");
        assert_eq!(describe("#install", &md), "#install");
    }

    // A long path gets line breaks after a separator, so the tooltip stays on
    // the screen; nothing is left out.
    #[test]
    fn test_long_targets_wrap_without_losing_text() {
        assert_eq!(wrap_for_display(r"C:\docs\b.md"), r"C:\docs\b.md");
        let long = format!(r"C:\{}\file.md", "a-rather-long-folder-name\\".repeat(12));
        let wrapped = wrap_for_display(&long);
        assert_eq!(wrapped.replace("\r\n", ""), long);
        for line in wrapped.split("\r\n") {
            assert!(line.chars().count() <= LINE_CHARS * 3 / 2, "{line}");
        }
        assert!(wrapped.split("\r\n").all(|l| l.ends_with('\\') || l.ends_with("file.md")), "{wrapped}");
        let no_separator = "x".repeat(400);
        assert_eq!(wrap_for_display(&no_separator).replace("\r\n", ""), no_separator);
        assert!(wrap_for_display(&no_separator).split("\r\n").all(|l| l.len() <= LINE_CHARS * 3 / 2));
    }

    #[test]
    fn test_links_to_a_heading_in_another_document_show_the_heading() {
        let md = LinkAction::OpenMarkdown(PathBuf::from(r"C:\docs\b.md"));
        assert_eq!(describe_url("https://mdview.example/docs/b.md#setup", "b.md#setup", &md), r"C:\docs\b.md#setup");
        assert_eq!(describe_url("https://mdview.example/docs/b.md#", "b.md#", &md), r"C:\docs\b.md");
        assert_eq!(describe_url("https://mdview.example/docs/a.md#top", "#top", &md), "#top");
        let long = format!(r"C:\{}\file.md", "folder\\".repeat(40));
        let action = LinkAction::OpenMarkdown(PathBuf::from(&long));
        assert_eq!(describe_url("x", "x", &action), long, "shown in full, never cut off");
    }
}
