#![allow(non_snake_case)]

mod markdown;
mod viewer;

use std::ffi::{c_char, c_int, CStr};
use std::ptr;
use widestring::U16CStr;
use windows::Win32::Foundation::HWND;

// WLX Plugin Constants
const LCP_DARKMODE: c_int = 128;

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
    // Read the markdown file
    let markdown_content = match std::fs::read_to_string(file_path) {
        Ok(content) => content,
        Err(_) => return HWND::default(),
    };

    // Convert to HTML
    let html_body = markdown::markdown_to_html(&markdown_content);
    let full_html = markdown::wrap_html(&html_body, dark_mode);

    // Create viewer window with WebView2
    match viewer::create_viewer(parent, &full_html) {
        Ok(hwnd) => hwnd,
        Err(_) => HWND::default(),
    }
}

// Additional optional exports for enhanced functionality

/// Search for text in the document
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchText(
    _list_win: HWND,
    _search_string: *const c_char,
    _search_parameter: c_int,
) -> c_int {
    // TODO: Implement search via WebView2 JavaScript
    0 // LISTPLUGIN_OK
}

/// Search for text (Unicode version)
#[unsafe(no_mangle)]
pub extern "system" fn ListSearchTextW(
    _list_win: HWND,
    _search_string: *const u16,
    _search_parameter: c_int,
) -> c_int {
    // TODO: Implement search via WebView2 JavaScript
    0
}

/// Handle commands (copy, select all, etc.)
#[unsafe(no_mangle)]
pub extern "system" fn ListSendCommand(
    list_win: HWND,
    command: c_int,
    _parameter: c_int,
) -> c_int {
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
