#![allow(non_snake_case)]

mod files;
mod links;
mod markdown;
mod settings;
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
    let file_path = unsafe {
        match CStr::from_ptr(file_to_load).to_str() {
            Ok(s) => s.to_string(),
            Err(_) => return HWND::default(),
        }
    };

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

// Search parameter flags
const LP_FORWARDONLY: c_int = 2;
const LP_WHOLEWORDS: c_int = 4;
const LP_MATCHCASE: c_int = 8;

/// Search for text in the document
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchText(
    list_win: HWND,
    search_string: *const c_char,
    search_parameter: c_int,
) -> c_int {
    let search_str = unsafe {
        match CStr::from_ptr(search_string).to_str() {
            Ok(s) => s.to_string(),
            Err(_) => return 1, // LISTPLUGIN_ERROR
        }
    };
    do_search(list_win, &search_str, search_parameter)
}

/// Search for text (Unicode version)
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchTextW(
    list_win: HWND,
    search_string: *const u16,
    search_parameter: c_int,
) -> c_int {
    let search_str = unsafe {
        match U16CStr::from_ptr_str(search_string).to_string() {
            Ok(s) => s,
            Err(_) => return 1, // LISTPLUGIN_ERROR
        }
    };
    do_search(list_win, &search_str, search_parameter)
}

fn do_search(list_win: HWND, search_string: &str, search_parameter: c_int) -> c_int {
    let case_sensitive = (search_parameter & LP_MATCHCASE) != 0;
    let whole_word = (search_parameter & LP_WHOLEWORDS) != 0;
    let forward_only = (search_parameter & LP_FORWARDONLY) != 0;

    // Escape quotes in search string for JavaScript
    let escaped = search_string.replace('\\', "\\\\").replace('"', "\\\"");

    // Use window.find() for search
    // window.find(string, caseSensitive, backwards, wrapAround, wholeWord, searchInFrames, showDialog)
    let script = format!(
        "window.find(\"{}\", {}, {}, {}, {}, false, false)",
        escaped,
        case_sensitive,
        false,         // backwards (we search forward)
        !forward_only, // wrapAround
        whole_word
    );

    viewer::execute_script(list_win, &script);
    0 // LISTPLUGIN_OK
}

/// Handle commands (copy, select all, etc.)
#[unsafe(no_mangle)]
pub extern "system" fn ListSendCommand(list_win: HWND, command: c_int, _parameter: c_int) -> c_int {
    const LC_COPY: c_int = 1;
    const LC_SELECTALL: c_int = 3;

    match command {
        LC_COPY => {
            viewer::execute_script(list_win, "document.execCommand('copy')");
            1 // LISTPLUGIN_OK
        }
        LC_SELECTALL => {
            viewer::execute_script(list_win, "document.execCommand('selectAll')");
            1
        }
        _ => 0,
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
}
