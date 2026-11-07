#![allow(non_snake_case)]

mod clipboard;
mod contextmenu;
mod files;
mod fonts;
mod links;
mod linkstatus;
mod markdown;
mod messages;
mod settings;
mod settings_page;
mod shortcuts;
mod viewer;
mod vpath;
mod zone;

use std::ffi::{CStr, c_char, c_int};
use std::ptr;
use widestring::U16CStr;
use windows::Win32::Foundation::HWND;

/// This crate is the Total Commander plugin (see settings::store).
pub(crate) const IS_PLUGIN: bool = true;

// WLX Plugin Constants
const LCP_DARKMODE: c_int = 128;

/// WLX SDK ListDefaultParamStruct. `default_ini_name` is ANSI (MAX_PATH).
#[repr(C)]
pub struct ListDefaultParamStruct {
    size: c_int,
    plugin_interface_version_low: u32,
    plugin_interface_version_hi: u32,
    default_ini_name: [c_char; 260],
}

/// Called by TC right after loading the plugin, before ListLoad (TC 5.51+).
/// `DefaultIniName` is in the folder of wincmd.ini; MDView keeps its own
/// mdview.ini there (#34).
#[unsafe(no_mangle)]
pub extern "system" fn ListSetDefaultParams(dps: *const ListDefaultParamStruct) {
    let store = match unsafe { dps.as_ref() } {
        None => settings::Store::Unavailable("Total Commander passed no settings".to_string()),
        Some(dps) => match ini_folder_from_default_ini_name(&dps.default_ini_name) {
            Ok(folder) => settings::ini_store_for(&folder),
            Err(reason) => settings::Store::Unavailable(reason),
        },
    };
    let in_use = settings::use_store(store);
    viewer::log_debug(&format!("settings: {in_use:?}"));
}

/// The folder of TC's `DefaultIniName` (ANSI, NUL-terminated, MAX_PATH).
fn ini_folder_from_default_ini_name(name: &[c_char; 260]) -> Result<std::path::PathBuf, String> {
    use windows::Win32::Globalization::{CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS, MultiByteToWideChar};
    let bytes: Vec<u8> = name.iter().map(|&c| c as u8).collect();
    let len = bytes
        .iter()
        .position(|&b| b == 0)
        .ok_or_else(|| "Total Commander's settings path is not terminated".to_string())?;
    if len == 0 {
        return Err("Total Commander reported an empty settings path".to_string());
    }
    let mut wide = vec![0u16; len * 2 + 1];
    let n = unsafe { MultiByteToWideChar(CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), &bytes[..len], Some(&mut wide)) };
    if n <= 0 {
        return Err("Total Commander's settings path could not be read".to_string());
    }
    let path = std::path::PathBuf::from(String::from_utf16_lossy(&wide[..n as usize]));
    path.parent()
        .map(|p| p.to_path_buf())
        .filter(|p| !p.as_os_str().is_empty())
        .ok_or_else(|| format!("Total Commander's settings path {} has no folder", path.display()))
}

/// Detection string for Total Commander - handles .md and .markdown files
#[unsafe(no_mangle)]
pub extern "system" fn ListGetDetectString(detect_string: *mut c_char, maxlen: c_int) {
    let detect = b"EXT=\"MD\" | EXT=\"MARKDOWN\"\0";
    let len = detect.len().min(maxlen as usize);
    unsafe {
        ptr::copy_nonoverlapping(detect.as_ptr(), detect_string as *mut u8, len);
    }
}

/// Load a file (ANSI version)
#[unsafe(no_mangle)]
pub extern "system" fn ListLoad(
    parent_win: HWND,
    file_to_load: *const c_char,
    show_flags: c_int,
) -> HWND {
    // TC passes the ANSI version's path in the Windows code page.
    let Some(file_path) = (unsafe { ansi_to_string(file_to_load) }) else { return HWND::default() };

    let dark_mode = (show_flags & LCP_DARKMODE) != 0;
    load_markdown_file(parent_win, &file_path, dark_mode)
}

/// Load a file (Unicode version)
#[unsafe(no_mangle)]
pub extern "system" fn ListLoadW(
    parent_win: HWND,
    file_to_load: *const u16,
    show_flags: c_int,
) -> HWND {
    let file_path = unsafe {
        match U16CStr::from_ptr_str(file_to_load).to_string() {
            Ok(s) => s,
            Err(_) => return HWND::default(),
        }
    };

    let dark_mode = (show_flags & LCP_DARKMODE) != 0;
    load_markdown_file(parent_win, &file_path, dark_mode)
}

/// Close the viewer window
#[unsafe(no_mangle)]
pub extern "system" fn ListCloseWindow(list_win: HWND) {
    viewer::close_window(list_win);
}

fn load_markdown_file(parent: HWND, file_path: &str, dark_mode: bool) -> HWND {
    // #55: only Markdown; declining lets TC use another viewer for the file.
    if !links::is_markdown(std::path::Path::new(file_path)) {
        return HWND::default();
    }
    // Only check that the file can be opened; TC shows its own message if not.
    // The content is read later, after its Mark of the Web was checked (#24).
    if files::open_read_only(std::path::Path::new(file_path)).is_err() {
        return HWND::default();
    }

    // Create viewer window with WebView2
    match viewer::create_viewer(parent, "", Some(file_path), dark_mode) {
        Ok(hwnd) => hwnd,
        Err(_) => HWND::default(),
    }
}

// Additional optional exports for enhanced functionality

// WLX SDK return values and flags (listplug.h).
const LISTPLUGIN_OK: c_int = 0;
const LISTPLUGIN_ERROR: c_int = 1;
const LCS_FINDFIRST: c_int = 1;
const LCS_MATCHCASE: c_int = 2;
const LCS_WHOLEWORDS: c_int = 4;
const LCS_BACKWARDS: c_int = 8;
const LC_COPY: c_int = 1;
const LC_SELECTALL: c_int = 3;

/// TC's search options for `ListSearchText` (#59).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SearchFlags {
    /// Start at the top (at the bottom when searching backwards), not after
    /// the current match.
    find_first: bool,
    match_case: bool,
    whole_words: bool,
    backwards: bool,
}

fn search_flags(parameter: c_int) -> SearchFlags {
    SearchFlags {
        find_first: parameter & LCS_FINDFIRST != 0,
        match_case: parameter & LCS_MATCHCASE != 0,
        whole_words: parameter & LCS_WHOLEWORDS != 0,
        backwards: parameter & LCS_BACKWARDS != 0,
    }
}

/// The page script for one search step; it returns whether a match was found
/// and selected. The text is a JSON string, so it can't end the script.
fn search_script(text: &str, flags: SearchFlags) -> String {
    let start = if flags.find_first {
        format!(
            "var s = window.getSelection(); s.removeAllRanges(); var r = document.createRange(); \
             r.selectNodeContents(document.body); r.collapse({}); s.addRange(r);",
            !flags.backwards
        )
    } else {
        String::new()
    };
    // window.find(text, caseSensitive, backwards, wrapAround, wholeWord, searchInFrames, showDialog)
    format!(
        "(function() {{ {start} return window.find({}, {}, {}, false, {}, false, false); }})()",
        markdown::script_string(text),
        flags.match_case,
        flags.backwards,
        flags.whole_words,
    )
}

/// Search for text in the document (TC's F7, then F3 / Shift+F3).
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchText(
    list_win: HWND,
    search_string: *const c_char,
    search_parameter: c_int,
) -> c_int {
    // TC passes the ANSI version text in the Windows code page.
    match unsafe { ansi_to_string(search_string) } {
        Some(text) => do_search(list_win, &text, search_parameter),
        None => LISTPLUGIN_ERROR,
    }
}

/// Search for text (Unicode version)
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchTextW(
    list_win: HWND,
    search_string: *const u16,
    search_parameter: c_int,
) -> c_int {
    if search_string.is_null() {
        return LISTPLUGIN_ERROR;
    }
    match unsafe { U16CStr::from_ptr_str(search_string) }.to_string() {
        Ok(text) => do_search(list_win, &text, search_parameter),
        Err(_) => LISTPLUGIN_ERROR,
    }
}

/// A NUL-terminated string in the Windows code page.
unsafe fn ansi_to_string(text: *const c_char) -> Option<String> {
    use windows::Win32::Globalization::{CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS, MultiByteToWideChar};
    if text.is_null() {
        return None;
    }
    let bytes = unsafe { CStr::from_ptr(text) }.to_bytes();
    if bytes.is_empty() {
        return Some(String::new());
    }
    let len = unsafe { MultiByteToWideChar(CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), bytes, None) };
    if len <= 0 {
        return None;
    }
    let mut wide = vec![0u16; len as usize];
    let written = unsafe { MultiByteToWideChar(CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), bytes, Some(&mut wide)) };
    (written == len).then(|| String::from_utf16_lossy(&wide))
}

/// One search step. Returns OK when a match was found and selected, else an
/// error, so TC says it was not found.
fn do_search(list_win: HWND, text: &str, search_parameter: c_int) -> c_int {
    if text.is_empty() {
        return LISTPLUGIN_ERROR;
    }
    let script = search_script(text, search_flags(search_parameter));
    match viewer::execute_script_sync(list_win, &script, std::time::Duration::from_secs(2)) {
        Some(result) if result == "true" => LISTPLUGIN_OK,
        _ => LISTPLUGIN_ERROR,
    }
}

/// Copy and select all from TC's menu.
#[unsafe(no_mangle)]
pub extern "system" fn ListSendCommand(list_win: HWND, command: c_int, _parameter: c_int) -> c_int {
    match command {
        LC_COPY => {
            viewer::execute_script(list_win, "document.execCommand('copy')");
            LISTPLUGIN_OK
        }
        LC_SELECTALL => {
            viewer::execute_script(list_win, "document.execCommand('selectAll')");
            LISTPLUGIN_OK
        }
        _ => LISTPLUGIN_ERROR,
    }
}

/// TC's print command: WebView2's print dialog (printer choice, preview).
#[unsafe(no_mangle)]
pub extern "system" fn ListPrint(
    list_win: HWND,
    _file_to_print: *const c_char,
    _def_printer: *const c_char,
    _print_flags: c_int,
    _margins: *const windows::Win32::Foundation::RECT,
) -> c_int {
    print(list_win)
}

#[unsafe(no_mangle)]
pub extern "system" fn ListPrintW(
    list_win: HWND,
    _file_to_print: *const u16,
    _def_printer: *const u16,
    _print_flags: c_int,
    _margins: *const windows::Win32::Foundation::RECT,
) -> c_int {
    print(list_win)
}

fn print(list_win: HWND) -> c_int {
    match viewer::show_print_ui(list_win) {
        Ok(()) => LISTPLUGIN_OK,
        Err(e) => {
            viewer::show_error(list_win, &format!("Could not open the print dialog:\n\n{e}"));
            LISTPLUGIN_ERROR
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ini_name(s: &[u8]) -> [c_char; 260] {
        let mut buf = [0 as c_char; 260];
        for (i, &b) in s.iter().enumerate() {
            buf[i] = b as c_char;
        }
        buf
    }

    // #34: MDView's settings go next to TC's DefaultIniName (wincmd.ini's folder).
    #[test]
    fn test_ini_folder_is_the_folder_of_default_ini_name() {
        let folder = ini_folder_from_default_ini_name(&ini_name(b"C:\\Users\\x\\AppData\\Roaming\\GHISLER\\lsplugin.ini")).unwrap();
        assert_eq!(folder, std::path::PathBuf::from("C:\\Users\\x\\AppData\\Roaming\\GHISLER"));
    }

    #[test]
    fn test_unusable_default_ini_names_are_rejected() {
        assert!(ini_folder_from_default_ini_name(&ini_name(b"")).unwrap_err().contains("empty"));
        assert!(ini_folder_from_default_ini_name(&[b'a' as c_char; 260]).unwrap_err().contains("not terminated"));
        assert!(ini_folder_from_default_ini_name(&ini_name(b"lsplugin.ini")).unwrap_err().contains("no folder"));
    }

    // #59: TC's search flags (listplug.h): before, 2 was read as "forward
    // only" and 8 as "match case".
    #[test]
    fn test_search_flags_follow_the_wlx_sdk() {
        let none = SearchFlags { find_first: false, match_case: false, whole_words: false, backwards: false };
        assert_eq!(search_flags(0), none);
        assert_eq!(search_flags(1), SearchFlags { find_first: true, ..none });
        assert_eq!(search_flags(2), SearchFlags { match_case: true, ..none });
        assert_eq!(search_flags(4), SearchFlags { whole_words: true, ..none });
        assert_eq!(search_flags(8), SearchFlags { backwards: true, ..none });
        assert_eq!(search_flags(15), SearchFlags { find_first: true, match_case: true, whole_words: true, backwards: true });
    }

    #[test]
    fn test_search_script_passes_the_options_and_the_text_safely() {
        let flags = search_flags(LCS_MATCHCASE | LCS_BACKWARDS);
        let script = search_script("a\"b</x>\\", flags);
        let expected = concat!("window.find(\"a\\\"b\\u", "003c/x\\u", "003e\\\\\", true, true, false, false, false, false)");
        assert!(script.contains(expected), "{script}");
        assert!(!script.contains("removeAllRanges"), "find next continues after the current match");
        let first = search_script("x", search_flags(LCS_FINDFIRST));
        assert!(first.contains("removeAllRanges") && first.contains("r.collapse(true)"), "starts at the top");
        let first_back = search_script("x", search_flags(LCS_FINDFIRST | LCS_BACKWARDS));
        assert!(first_back.contains("r.collapse(false)"), "backwards starts at the bottom");
    }

    #[test]
    fn test_ansi_search_text_uses_the_windows_code_page() {
        let text = std::ffi::CString::new("plain text").unwrap();
        assert_eq!(unsafe { ansi_to_string(text.as_ptr()) }.as_deref(), Some("plain text"));
        assert_eq!(unsafe { ansi_to_string(std::ptr::null()) }, None);
    }
}
