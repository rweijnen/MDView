#![windows_subsystem = "windows"]

mod dark_menu;
mod files;
mod links;
mod markdown;
mod settings;
mod terminal;
mod vpath;
mod zone;

use std::cell::RefCell;
use std::env;
use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;
use std::rc::Rc;

fn print_usage_console() {
    let usage = format!(
        "MDView - Markdown Viewer v{}\n\
         Copyright 2026 Remko Weijnen - Mozilla Public License 2.0\n\
         https://github.com/rweijnen/MDView\n\n\
         Usage: mdview [OPTIONS] [FILE]\n\n\
         Options:\n\
         \x20 --gui        Open in GUI window\n\
         \x20 --term       Output with terminal colors/formatting\n\
         \x20 --html       Output full HTML document to stdout\n\
         \x20 --body       Output HTML body only (no wrapper)\n\
         \x20 --text       Output plain text (no formatting)\n\
         \x20 --register   Register as .md file viewer (Open With)\n\
         \x20 --unregister Remove .md file viewer registration\n\
         \x20 -h, --help   Show this help message\n\n\
         If no FILE is specified, reads from stdin (CLI mode only).\n\n\
         Examples:\n\
         \x20 mdview README.md              # Open in GUI window\n\
         \x20 mdview --term README.md       # Output with terminal colors\n\
         \x20 cat doc.md | mdview           # Piped input, terminal output\n\
         \x20 mdview --html README.md       # Output HTML to stdout\n",
        env!("CARGO_PKG_VERSION")
    );
    write_console(&usage);
}

#[derive(Default)]
struct Options {
    gui_mode: bool,
    terminal_mode: bool,
    html_full: bool,
    html_body: bool,
    plain_text: bool,
    register: bool,
    unregister: bool,
    file_path: Option<String>,
}

fn parse_args(has_console: bool) -> Result<Options, String> {
    let args: Vec<String> = env::args().skip(1).collect();
    let mut opts = Options::default();

    for arg in args {
        match arg.as_str() {
            "-h" | "--help" => {
                print_usage_console();
                std::process::exit(0);
            }
            "--gui" => opts.gui_mode = true,
            "--term" | "--terminal" => opts.terminal_mode = true,
            "--html" => opts.html_full = true,
            "--body" => opts.html_body = true,
            "--text" => opts.plain_text = true,
            "--register" => opts.register = true,
            "--unregister" => opts.unregister = true,
            s if s.starts_with('-') => {
                return Err(format!("Unknown option: {}", s));
            }
            path => {
                if opts.file_path.is_some() {
                    return Err("Multiple input files not supported".to_string());
                }
                opts.file_path = Some(path.to_string());
            }
        }
    }

    // Register/unregister are standalone operations
    if opts.register || opts.unregister {
        if opts.register && opts.unregister {
            return Err("--register and --unregister are mutually exclusive".to_string());
        }
        let other_flags = opts.html_full as u8
            + opts.html_body as u8
            + opts.plain_text as u8
            + opts.terminal_mode as u8
            + opts.gui_mode as u8;
        if other_flags > 0 || opts.file_path.is_some() {
            return Err(
                "--register and --unregister cannot be combined with other options".to_string(),
            );
        }
        return Ok(opts);
    }

    // Validate mutually exclusive options
    let cli_format_count = opts.html_full as u8
        + opts.html_body as u8
        + opts.plain_text as u8
        + opts.terminal_mode as u8;
    if cli_format_count > 1 {
        return Err(
            "Options --term, --html, --body, and --text are mutually exclusive".to_string(),
        );
    }

    // Default behavior based on whether we have a console (terminal) or not (double-clicked)
    if !opts.gui_mode && cli_format_count == 0 {
        if has_console {
            // Launched from terminal: use terminal output mode
            opts.terminal_mode = true;
        } else {
            // Double-clicked / no console: use GUI mode
            opts.gui_mode = true;
        }
    }

    Ok(opts)
}

/// Detect if Windows is using dark mode (apps theme)
fn is_windows_dark_mode() -> bool {
    use windows::Win32::System::Registry::{
        HKEY_CURRENT_USER, KEY_READ, REG_VALUE_TYPE, RegCloseKey, RegOpenKeyExW, RegQueryValueExW,
    };
    use windows::core::PCWSTR;

    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> =
            "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\0"
                .encode_utf16()
                .collect();

        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            KEY_READ,
            &mut hkey,
        )
        .is_ok()
        {
            let value_name: Vec<u16> = "AppsUseLightTheme\0".encode_utf16().collect();
            let mut data: u32 = 1;
            let mut data_size: u32 = std::mem::size_of::<u32>() as u32;
            let mut data_type: REG_VALUE_TYPE = REG_VALUE_TYPE::default();

            let result = RegQueryValueExW(
                hkey,
                PCWSTR(value_name.as_ptr()),
                None,
                Some(&mut data_type),
                Some(&mut data as *mut u32 as *mut u8),
                Some(&mut data_size),
            );

            let _ = RegCloseKey(hkey);

            if result.is_ok() {
                return data == 0; // 0 = dark mode, 1 = light mode
            }
        }
    }
    false // Default to light mode
}

/// Try to attach to parent process console for CLI output
/// Returns true if successfully attached (meaning we were launched from a terminal)
fn attach_console() -> bool {
    use windows::Win32::System::Console::{ATTACH_PARENT_PROCESS, AttachConsole};

    unsafe { AttachConsole(ATTACH_PARENT_PROCESS).is_ok() }
}

/// Write text directly to console and print newline to "complete" the prompt line
fn write_console(text: &str) {
    use windows::Win32::System::Console::{GetStdHandle, STD_OUTPUT_HANDLE, WriteConsoleW};

    unsafe {
        if let Ok(handle) = GetStdHandle(STD_OUTPUT_HANDLE) {
            // Print newline first to move past the prompt that cmd already showed
            let newline: Vec<u16> = "\r\n".encode_utf16().collect();
            let mut written = 0u32;
            let _ = WriteConsoleW(handle, &newline, Some(&mut written), None);

            // Now print our text
            let wide: Vec<u16> = text.encode_utf16().collect();
            let _ = WriteConsoleW(handle, &wide, Some(&mut written), None);
        }
    }
}

/// Enable ANSI/Virtual Terminal Processing on Windows console
/// This allows legacy cmd.exe to process ANSI escape codes
fn enable_virtual_terminal_processing() {
    use windows::Win32::System::Console::{
        CONSOLE_MODE, ENABLE_VIRTUAL_TERMINAL_PROCESSING, GetConsoleMode, GetStdHandle,
        STD_OUTPUT_HANDLE, SetConsoleMode,
    };

    unsafe {
        if let Ok(handle) = GetStdHandle(STD_OUTPUT_HANDLE) {
            let mut mode = CONSOLE_MODE::default();
            if GetConsoleMode(handle, &mut mode).is_ok() {
                let _ = SetConsoleMode(handle, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
            }
        }
    }
}

/// Send Enter key to release the command prompt after printing output.
/// Only sends if the console window is in the foreground (safety check to avoid
/// sending Enter to other apps if running in background).
fn send_enter_key() {
    use windows::Win32::System::Console::GetConsoleWindow;
    use windows::Win32::UI::Input::KeyboardAndMouse::{
        INPUT, INPUT_KEYBOARD, KEYBDINPUT, KEYEVENTF_KEYUP, SendInput, VK_RETURN,
    };
    use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;

    unsafe {
        // Only send Enter if console window is in focus
        let console_hwnd = GetConsoleWindow();
        let foreground_hwnd = GetForegroundWindow();

        if console_hwnd.0 != std::ptr::null_mut() && console_hwnd == foreground_hwnd {
            let mut inputs: [INPUT; 2] = std::mem::zeroed();

            // Key down
            inputs[0].r#type = INPUT_KEYBOARD;
            inputs[0].Anonymous.ki = KEYBDINPUT {
                wVk: VK_RETURN,
                wScan: 0,
                dwFlags: Default::default(),
                time: 0,
                dwExtraInfo: 0,
            };

            // Key up
            inputs[1].r#type = INPUT_KEYBOARD;
            inputs[1].Anonymous.ki = KEYBDINPUT {
                wVk: VK_RETURN,
                wScan: 0,
                dwFlags: KEYEVENTF_KEYUP,
                time: 0,
                dwExtraInfo: 0,
            };

            SendInput(&inputs, std::mem::size_of::<INPUT>() as i32);
        }
    }
}

fn read_input(file_path: Option<&str>) -> io::Result<String> {
    match file_path {
        Some(path) => {
            let path = Path::new(path);
            if !path.exists() {
                return Err(io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("File not found: {}", path.display()),
                ));
            }
            files::read_text(path).map_err(|e| io::Error::other(format!("{}: {e}", path.display())))
        }
        None => {
            let mut buffer = String::new();
            io::stdin().read_to_string(&mut buffer)?;
            Ok(buffer)
        }
    }
}

fn main() {
    // Try to attach to parent console - this tells us if we're launched from a terminal
    let has_console = attach_console();

    let opts = match parse_args(has_console) {
        Ok(o) => o,
        Err(e) => {
            eprintln!("Error: {}", e);
            eprintln!("Use --help for usage information.");
            std::process::exit(1);
        }
    };

    // Handle --register / --unregister
    if opts.register {
        match register_file_association() {
            Ok(()) => {
                write_console(concat!(app_name!(), " file association registered successfully."));
                unsafe {
                    let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
                    open_association_settings(None);
                    CoUninitialize();
                }
                send_enter_key();
                std::process::exit(0);
            }
            Err(e) => {
                write_console(&format!("Error: {}", e));
                send_enter_key();
                std::process::exit(1);
            }
        }
    }

    if opts.unregister {
        match unregister_file_association() {
            Ok(()) => {
                write_console(concat!(app_name!(), " file association removed successfully."));
                send_enter_key();
                std::process::exit(0);
            }
            Err(e) => {
                write_console(&format!("Error: {}", e));
                send_enter_key();
                std::process::exit(1);
            }
        }
    }

    // If no file and running in CLI/terminal mode, show help and exit
    if opts.file_path.is_none() && !opts.gui_mode {
        print_usage_console();
        send_enter_key(); // Release command prompt
        std::process::exit(0);
    }

    if opts.gui_mode {
        // GUI mode - open window with WebView2
        let (title, full_html) = if let Some(ref path) = opts.file_path {
            // The document is not read here: it is loaded through its virtual
            // URL and served by create_virtual_resource_response, which first
            // checks its Mark of the Web (#24).
            if !Path::new(path).exists() {
                eprintln!("Error reading input: File not found: {path}");
                std::process::exit(1);
            }
            let full_html = String::new();
            let title = Path::new(path)
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            (title, full_html)
        } else {
            // No file - show welcome screen
            let dark_mode = is_windows_dark_mode();
            let welcome_bars = markdown::document_bars(&markdown::PageBars {
                remote_images: markdown::RemoteImages::Blocked,
                ask_to_register: ask_to_register(),
                ask_network_files: false,
                closed: &[],
            });
            let welcome_html = markdown::wrap_html(
                "<div style=\"text-align: center; margin-top: 100px; color: #888;\">\
                 <h1>MDView</h1>\
                 <p>Open a Markdown file using <strong>File &gt; Open</strong> (Ctrl+O)</p>\
                 <p>or drag and drop a .md file onto this window.</p>\
                 </div>",
                dark_mode,
                markdown::RemoteImages::Blocked,
                &welcome_bars,
            );
            PAGE_BARS.with(|b| *b.borrow_mut() = welcome_bars);
            (APP_NAME.to_string(), welcome_html)
        };

        if let Err(e) = run_gui(&title, &full_html, opts.file_path.as_deref()) {
            eprintln!("Error: {}", e);
            std::process::exit(1);
        }
    } else {
        // CLI mode needs content
        let markdown_content = match read_input(opts.file_path.as_deref()) {
            Ok(content) => content,
            Err(e) => {
                eprintln!("Error reading input: {}", e);
                std::process::exit(1);
            }
        };
        // CLI mode - output to stdout
        let output = if opts.terminal_mode {
            // Enable ANSI processing on Windows console
            enable_virtual_terminal_processing();
            let caps = terminal::TerminalCaps::detect();
            terminal::render_to_terminal(&markdown_content, &caps)
        } else if opts.plain_text {
            markdown::markdown_to_plain_text(&markdown_content)
        } else if opts.html_body {
            markdown::markdown_to_html(&markdown_content)
        } else {
            let html_body = markdown::markdown_to_html(&markdown_content);
            let dark_mode = is_windows_dark_mode();
            // An exported page is the user's own file, opened in their browser:
            // no banner works there, so remote images are allowed.
            markdown::wrap_html(&html_body, dark_mode, markdown::RemoteImages::Allowed, &[])
        };

        if let Err(e) = io::stdout().write_all(output.as_bytes()) {
            eprintln!("Error writing output: {}", e);
            std::process::exit(1);
        }
    }
}

// ============================================================================
// GUI Implementation using WebView2
// ============================================================================

use webview2_com::Microsoft::Web::WebView2::Win32::*;
use webview2_com::{
    CreateCoreWebView2ControllerCompletedHandler, CreateCoreWebView2EnvironmentCompletedHandler,
    pwstr_from_str,
};
use widestring::U16CString;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Dwm::{
    DWMWA_USE_IMMERSIVE_DARK_MODE, DWMWA_WINDOW_CORNER_PREFERENCE, DWMWCP_ROUND,
    DwmSetWindowAttribute,
};
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{
    CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx,
    CoUninitialize,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Registry::{
    HKEY_CURRENT_USER, KEY_READ, KEY_WRITE, REG_CREATE_KEY_DISPOSITION, REG_DWORD, REG_SZ,
    REG_VALUE_TYPE, RegCloseKey, RegCreateKeyExW, RegDeleteTreeW, RegDeleteValueW, RegOpenKeyExW,
    RegQueryValueExW, RegSetValueExW,
};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::Shell::{
    DragFinish, DragQueryFileW, FOS_FILEMUSTEXIST, FOS_PATHMUSTEXIST, FileOpenDialog, HDROP,
    IFileOpenDialog, SHCNE_ASSOCCHANGED, SHCNF_IDLIST,
    SHChangeNotify, SHCreateMemStream, SIGDN_FILESYSPATH, ShellExecuteW,
};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::Interface;
use windows::core::PCWSTR;

const WINDOW_CLASS: &str = "MDViewWindow";

// Dark theme color for client area background
const DARK_CLIENT: u32 = 0x00202020;

// Menu item IDs
const IDM_FILE_OPEN: u32 = 1001;
const IDM_FILE_EXIT: u32 = 1002;
const IDM_HELP_ABOUT: u32 = 2001;
const IDM_EDIT_FIND: u32 = 1201;
const IDM_FILE_REGISTER: u32 = 1010;
const IDM_FILE_UNREGISTER: u32 = 1011;
const IDM_FILE_RECENT_BASE: u32 = 1100; // 1100-1109 for recent files

// The product name is defined once, in settings::app_name!().
use settings::{APP_NAME, REGISTRY_KEY, app_name};

/// This crate is the standalone viewer, not the TC plugin (see settings::store).
pub(crate) const IS_PLUGIN: bool = false;

// File association registration
const PROGID: &str = concat!(app_name!(), ".Markdown");
const PROGID_DESCRIPTION: &str = "Markdown Document";
/// Name under which MDView is listed in HKCU\Software\RegisteredApplications;
/// the Settings deep link must use the same name.
const REGISTERED_APP_NAME: &str = APP_NAME;
const CAPABILITIES_KEY: &str = concat!("Software\\", app_name!(), "\\Capabilities");
const FILE_ASSOCIATIONS_KEY: &str =
    concat!("Software\\", app_name!(), "\\Capabilities\\FileAssociations");
const APP_DESCRIPTION: &str = "Fast, lightweight Markdown viewer";

// Undocumented uxtheme.dll ordinals for dark mode APIs
const UXTHEME_ORDINAL_ALLOW_DARK_MODE_FOR_WINDOW: usize = 133;
const UXTHEME_ORDINAL_SET_PREFERRED_APP_MODE: usize = 135;
const UXTHEME_ORDINAL_FLUSH_MENU_THEMES: usize = 136;

// PreferredAppMode values for SetPreferredAppMode
#[allow(dead_code)]
const PREFERRED_APP_MODE_DEFAULT: i32 = 0;
#[allow(dead_code)]
const PREFERRED_APP_MODE_ALLOW_DARK: i32 = 1;
const PREFERRED_APP_MODE_FORCE_DARK: i32 = 2;
#[allow(dead_code)]
const PREFERRED_APP_MODE_FORCE_LIGHT: i32 = 3;

/// Enable dark mode for menus using undocumented uxtheme.dll API
fn set_preferred_app_mode_dark() {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

    unsafe {
        let uxtheme: Vec<u16> = "uxtheme.dll\0".encode_utf16().collect();
        if let Ok(hmodule) = LoadLibraryW(PCWSTR(uxtheme.as_ptr())) {
            if let Some(set_preferred_app_mode) = GetProcAddress(
                hmodule,
                windows::core::PCSTR(UXTHEME_ORDINAL_SET_PREFERRED_APP_MODE as *const u8),
            ) {
                let func: extern "system" fn(i32) -> i32 =
                    std::mem::transmute(set_preferred_app_mode);
                func(PREFERRED_APP_MODE_FORCE_DARK);
            }
            // Flush menu themes to apply the change
            if let Some(flush_menu_themes) = GetProcAddress(
                hmodule,
                windows::core::PCSTR(UXTHEME_ORDINAL_FLUSH_MENU_THEMES as *const u8),
            ) {
                let func: extern "system" fn() = std::mem::transmute(flush_menu_themes);
                func();
            }
        }
    }
}

/// Allow dark mode for a specific window (needed for dark popup menus)
fn allow_dark_mode_for_window(hwnd: HWND) {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

    unsafe {
        let uxtheme: Vec<u16> = "uxtheme.dll\0".encode_utf16().collect();
        if let Ok(hmodule) = LoadLibraryW(PCWSTR(uxtheme.as_ptr())) {
            if let Some(allow_dark_mode) = GetProcAddress(
                hmodule,
                windows::core::PCSTR(UXTHEME_ORDINAL_ALLOW_DARK_MODE_FOR_WINDOW as *const u8),
            ) {
                let func: extern "system" fn(isize, bool) -> bool =
                    std::mem::transmute(allow_dark_mode);
                func(hwnd.0 as isize, true);
            }
        }
    }
}

/// Helper to insert a native menu item at a position
fn insert_menu_item(
    menu: HMENU,
    position: u32,
    id: u32,
    text: &str,
    is_grayed: bool,
) -> windows::core::Result<()> {
    unsafe {
        let text_wide: Vec<u16> = format!("{}\0", text).encode_utf16().collect();
        let flags = if is_grayed {
            MF_BYPOSITION | MF_STRING | MF_GRAYED
        } else {
            MF_BYPOSITION | MF_STRING
        };
        InsertMenuW(
            menu,
            position,
            flags,
            id as usize,
            PCWSTR(text_wide.as_ptr()),
        )?;
        Ok(())
    }
}

/// Create the application menu bar
fn create_menu() -> windows::core::Result<HMENU> {
    unsafe {
        let menu_bar = CreateMenu()?;
        let file_menu = CreatePopupMenu()?;
        let edit_menu = CreatePopupMenu()?;
        let help_menu = CreatePopupMenu()?;

        // File menu items (native Windows menus for proper dark mode styling)
        let open_text: Vec<u16> = "&Open\tCtrl+O\0".encode_utf16().collect();
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_FILE_OPEN as usize,
            PCWSTR(open_text.as_ptr()),
        )?;
        AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null())?;
        let recent_text: Vec<u16> = "Recent Files\0".encode_utf16().collect();
        AppendMenuW(
            file_menu,
            MF_STRING | MF_GRAYED,
            0,
            PCWSTR(recent_text.as_ptr()),
        )?;
        AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null())?;
        let register_text: Vec<u16> = "Register as .md Viewer...\0".encode_utf16().collect();
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_FILE_REGISTER as usize,
            PCWSTR(register_text.as_ptr()),
        )?;
        let unregister_text: Vec<u16> = "Unregister as .md Viewer\0".encode_utf16().collect();
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_FILE_UNREGISTER as usize,
            PCWSTR(unregister_text.as_ptr()),
        )?;
        AppendMenuW(file_menu, MF_SEPARATOR, 0, PCWSTR::null())?;
        let exit_text: Vec<u16> = "E&xit\0".encode_utf16().collect();
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_FILE_EXIT as usize,
            PCWSTR(exit_text.as_ptr()),
        )?;

        // Edit menu items
        let find_text: Vec<u16> = "&Find...\tCtrl+F\0".encode_utf16().collect();
        AppendMenuW(
            edit_menu,
            MF_STRING,
            IDM_EDIT_FIND as usize,
            PCWSTR(find_text.as_ptr()),
        )?;

        // Help menu items (native Windows menus)
        let about_text: Vec<u16> = concat!("&About ", app_name!(), "\0").encode_utf16().collect();
        AppendMenuW(
            help_menu,
            MF_STRING,
            IDM_HELP_ABOUT as usize,
            PCWSTR(about_text.as_ptr()),
        )?;

        // Add submenus to menu bar
        let file_text: Vec<u16> = "&File\0".encode_utf16().collect();
        AppendMenuW(
            menu_bar,
            MF_POPUP,
            file_menu.0 as usize,
            PCWSTR(file_text.as_ptr()),
        )?;

        let edit_text: Vec<u16> = "&Edit\0".encode_utf16().collect();
        AppendMenuW(
            menu_bar,
            MF_POPUP,
            edit_menu.0 as usize,
            PCWSTR(edit_text.as_ptr()),
        )?;

        let help_text: Vec<u16> = "&Help\0".encode_utf16().collect();
        AppendMenuW(
            menu_bar,
            MF_POPUP,
            help_menu.0 as usize,
            PCWSTR(help_text.as_ptr()),
        )?;

        Ok(menu_bar)
    }
}

/// Create accelerator table for keyboard shortcuts
fn create_accelerators() -> windows::core::Result<HACCEL> {
    unsafe {
        // These only fire while MDView's own window has focus; inside the
        // document, WebView2 handles Ctrl+F itself with the same find bar.
        let accels = [
            ACCEL {
                fVirt: FVIRTKEY | FCONTROL,
                key: 'O' as u16,
                cmd: IDM_FILE_OPEN as u16,
            },
            ACCEL {
                fVirt: FVIRTKEY | FCONTROL,
                key: 'F' as u16,
                cmd: IDM_EDIT_FIND as u16,
            },
        ];
        let haccel = CreateAcceleratorTableW(&accels)?;
        Ok(haccel)
    }
}

/// Show the modern IFileOpenDialog
fn show_open_file_dialog(hwnd: HWND) -> Option<String> {
    unsafe {
        let dialog: IFileOpenDialog =
            CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER).ok()?;

        // Set file type filters using raw COMDLG_FILTERSPEC
        let md_name: Vec<u16> = "Markdown Files\0".encode_utf16().collect();
        let md_spec: Vec<u16> = "*.md;*.markdown\0".encode_utf16().collect();
        let all_name: Vec<u16> = "All Files\0".encode_utf16().collect();
        let all_spec: Vec<u16> = "*.*\0".encode_utf16().collect();

        let filters = [
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(md_name.as_ptr()),
                pszSpec: PCWSTR(md_spec.as_ptr()),
            },
            COMDLG_FILTERSPEC {
                pszName: PCWSTR(all_name.as_ptr()),
                pszSpec: PCWSTR(all_spec.as_ptr()),
            },
        ];
        dialog.SetFileTypes(&filters).ok()?;

        // Set options
        let options = dialog.GetOptions().ok()?;
        dialog
            .SetOptions(options | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST)
            .ok()?;

        // Show dialog
        if dialog.Show(Some(hwnd)).is_err() {
            return None; // User cancelled
        }

        // Get result
        let result = dialog.GetResult().ok()?;
        let path = result.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
        let path_str = path.to_string().ok()?;
        windows::Win32::System::Com::CoTaskMemFree(Some(path.0 as *const _));
        Some(path_str)
    }
}

/// Show the About dialog
fn show_about_dialog(hwnd: HWND) {
    unsafe {
        let title: Vec<u16> = concat!("About ", app_name!(), "\0").encode_utf16().collect();
        let message = format!(
            "MDView - Markdown Viewer v{}\n\n\
             Copyright 2026 Remko Weijnen\n\
             Mozilla Public License 2.0\n\n\
             https://github.com/rweijnen/MDView\0",
            env!("CARGO_PKG_VERSION")
        );
        let message_wide: Vec<u16> = message.encode_utf16().collect();
        MessageBoxW(
            Some(hwnd),
            PCWSTR(message_wide.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONINFORMATION,
        );
    }
}

/// Load recent files list from registry
fn load_recent_files() -> Vec<String> {
    let mut files = Vec::new();
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("{}\\RecentFiles\0", REGISTRY_KEY)
            .encode_utf16()
            .collect();

        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            KEY_READ,
            &mut hkey,
        )
        .is_ok()
        {
            for i in 0..10 {
                let value_name: Vec<u16> = format!("File{}\0", i).encode_utf16().collect();
                let mut data = vec![0u16; 1024];
                let mut data_size = (data.len() * 2) as u32;
                let mut data_type = REG_VALUE_TYPE::default();

                if RegQueryValueExW(
                    hkey,
                    PCWSTR(value_name.as_ptr()),
                    None,
                    Some(&mut data_type),
                    Some(data.as_mut_ptr() as *mut u8),
                    Some(&mut data_size),
                )
                .is_ok()
                {
                    // Find null terminator
                    let len = data.iter().position(|&c| c == 0).unwrap_or(data.len());
                    if len > 0 {
                        let path = String::from_utf16_lossy(&data[..len]);
                        if !path.is_empty() && Path::new(&path).exists() {
                            files.push(path);
                        }
                    }
                }
            }
            let _ = RegCloseKey(hkey);
        }
    }
    files
}

/// Save recent files list to registry
fn save_recent_files(files: &[String]) {
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("{}\\RecentFiles\0", REGISTRY_KEY)
            .encode_utf16()
            .collect();
        let mut disposition = REG_CREATE_KEY_DISPOSITION::default();

        if RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            None,
            windows::Win32::System::Registry::REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut hkey,
            Some(&mut disposition),
        )
        .is_ok()
        {
            for i in 0..10 {
                let value_name: Vec<u16> = format!("File{}\0", i).encode_utf16().collect();
                if i < files.len() {
                    let data: Vec<u16> = format!("{}\0", files[i]).encode_utf16().collect();
                    let _ = RegSetValueExW(
                        hkey,
                        PCWSTR(value_name.as_ptr()),
                        Some(0),
                        REG_SZ,
                        Some(std::slice::from_raw_parts(
                            data.as_ptr() as *const u8,
                            data.len() * 2,
                        )),
                    );
                }
            }
            let _ = RegCloseKey(hkey);
        }
    }
}

/// Add a file to the recent files list
fn add_to_recent_files(file_path: &str) {
    RECENT_FILES.with(|files| {
        let mut files = files.borrow_mut();
        // Remove if already exists
        files.retain(|f| f != file_path);
        // Add to front
        files.insert(0, file_path.to_string());
        // Limit to 10
        files.truncate(10);
        save_recent_files(&files);
    });
    update_recent_files_menu();
}

/// Update the recent files submenu
fn update_recent_files_menu() {
    unsafe {
        MENU_HANDLE.with(|menu| {
            if let Some(menu) = menu.borrow().as_ref() {
                // Get the File menu (first popup in the menu bar)
                let file_menu = GetSubMenu(*menu, 0);
                if file_menu.is_invalid() {
                    return;
                }

                // Remove old recent file items (IDs 1100-1109)
                for id in IDM_FILE_RECENT_BASE..IDM_FILE_RECENT_BASE + 10 {
                    let _ = RemoveMenu(file_menu, id, MF_BYCOMMAND);
                }

                // Also remove the "Recent Files" placeholder if it exists
                let _ = RemoveMenu(file_menu, 2, MF_BYPOSITION);

                RECENT_FILES.with(|files| {
                    let files = files.borrow();
                    if files.is_empty() {
                        // Show disabled placeholder (native menu item)
                        let _ = insert_menu_item(file_menu, 2, 0, "(No Recent Files)", true);
                    } else {
                        // Add recent files (native menu items)
                        for (i, path) in files.iter().enumerate() {
                            let filename = Path::new(path)
                                .file_name()
                                .map(|n| n.to_string_lossy().to_string())
                                .unwrap_or_else(|| path.clone());
                            let text = format!("&{} {}", i + 1, filename);
                            let _ = insert_menu_item(
                                file_menu,
                                2 + i as u32,
                                IDM_FILE_RECENT_BASE + i as u32,
                                &text,
                                false,
                            );
                        }
                    }
                });
            }
        });
    }
}

/// Save window position to registry
/// The show state to remember for a window that is being closed. A window
/// closed while minimised (e.g. from the taskbar) is remembered in the state
/// it would be restored to, so the next start is not minimised.
fn show_cmd_to_save(placement: &WINDOWPLACEMENT) -> u32 {
    if placement.showCmd == SW_SHOWMINIMIZED.0 as u32 {
        if (placement.flags.0 & WPF_RESTORETOMAXIMIZED.0) != 0 {
            SW_SHOWMAXIMIZED.0 as u32
        } else {
            SW_SHOWNORMAL.0 as u32
        }
    } else {
        placement.showCmd
    }
}

/// The show state to start with: maximised if that was remembered, otherwise
/// normal. Never minimised, also for values saved by older versions.
fn show_cmd_to_restore(saved: i32) -> u32 {
    if saved == SW_SHOWMAXIMIZED.0 {
        SW_SHOWMAXIMIZED.0 as u32
    } else {
        SW_SHOWNORMAL.0 as u32
    }
}

fn save_window_position(hwnd: HWND) {
    unsafe {
        let mut placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            ..Default::default()
        };
        if GetWindowPlacement(hwnd, &mut placement).is_ok() {
            let mut hkey = std::mem::zeroed();
            let subkey: Vec<u16> = format!("{}\0", REGISTRY_KEY).encode_utf16().collect();
            let mut disposition = REG_CREATE_KEY_DISPOSITION::default();

            if RegCreateKeyExW(
                HKEY_CURRENT_USER,
                PCWSTR(subkey.as_ptr()),
                Some(0),
                None,
                windows::Win32::System::Registry::REG_OPTION_NON_VOLATILE,
                KEY_WRITE,
                None,
                &mut hkey,
                Some(&mut disposition),
            )
            .is_ok()
            {
                let values = [
                    ("Left", placement.rcNormalPosition.left as u32),
                    ("Top", placement.rcNormalPosition.top as u32),
                    ("Right", placement.rcNormalPosition.right as u32),
                    ("Bottom", placement.rcNormalPosition.bottom as u32),
                    ("ShowCmd", show_cmd_to_save(&placement)),
                ];
                for (name, value) in values {
                    let name_wide: Vec<u16> = format!("{}\0", name).encode_utf16().collect();
                    let _ = RegSetValueExW(
                        hkey,
                        PCWSTR(name_wide.as_ptr()),
                        Some(0),
                        REG_DWORD,
                        Some(std::slice::from_raw_parts(
                            &value as *const u32 as *const u8,
                            4,
                        )),
                    );
                }
                let _ = RegCloseKey(hkey);
            }
        }
    }
}

/// Restore window position from registry
fn restore_window_position(hwnd: HWND) -> bool {
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("{}\0", REGISTRY_KEY).encode_utf16().collect();

        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            KEY_READ,
            &mut hkey,
        )
        .is_err()
        {
            return false;
        }

        let read_dword = |name: &str| -> Option<i32> {
            let name_wide: Vec<u16> = format!("{}\0", name).encode_utf16().collect();
            let mut value: u32 = 0;
            let mut size: u32 = 4;
            let mut data_type = REG_VALUE_TYPE::default();
            if RegQueryValueExW(
                hkey,
                PCWSTR(name_wide.as_ptr()),
                None,
                Some(&mut data_type),
                Some(&mut value as *mut u32 as *mut u8),
                Some(&mut size),
            )
            .is_ok()
            {
                Some(value as i32)
            } else {
                None
            }
        };

        let left = read_dword("Left");
        let top = read_dword("Top");
        let right = read_dword("Right");
        let bottom = read_dword("Bottom");
        let show_cmd = read_dword("ShowCmd");

        let _ = RegCloseKey(hkey);

        if let (Some(l), Some(t), Some(r), Some(b), Some(cmd)) =
            (left, top, right, bottom, show_cmd)
        {
            // Validate window is at least partially visible
            if r > l && b > t && (r - l) >= 100 && (b - t) >= 100 {
                let placement = WINDOWPLACEMENT {
                    length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                    showCmd: show_cmd_to_restore(cmd),
                    rcNormalPosition: RECT {
                        left: l,
                        top: t,
                        right: r,
                        bottom: b,
                    },
                    ..Default::default()
                };
                let _ = SetWindowPlacement(hwnd, &placement);
                return true;
            }
        }
        false
    }
}

/// The URL a document is shown at (see `vpath`).
fn virtual_url_for_file(path: &Path) -> Option<String> {
    vpath::document_url(path)
}

fn content_type_for_path(path: &Path) -> &'static str {
    mime_guess::from_path(path)
        .first_raw()
        .unwrap_or("application/octet-stream")
}

/// Headers for files other than the Markdown page (images, and anything else a
/// document refers to): never run as a page with script, never sniffed into
/// another type. A local .html or .svg must not run script on MDView's origin.
const PASSIVE_FILE_HEADERS: &str = "\r\nContent-Security-Policy: default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'; sandbox\r\nX-Content-Type-Options: nosniff";

fn create_web_resource_response(
    environment: &ICoreWebView2Environment,
    bytes: &[u8],
    status_code: i32,
    reason: &str,
    content_type: &str,
    extra_headers: &str,
) -> Option<ICoreWebView2WebResourceResponse> {
    let stream = unsafe { SHCreateMemStream(Some(bytes)) }?;
    let reason_wide = pwstr_from_str(reason);
    let headers = format!(
        "Content-Type: {}\r\nAccess-Control-Allow-Origin: *{}",
        content_type, extra_headers
    );
    let headers_wide = pwstr_from_str(&headers);
    unsafe {
        environment
            .CreateWebResourceResponse(
                &stream,
                status_code,
                PCWSTR(reason_wide.as_ptr()),
                PCWSTR(headers_wide.as_ptr()),
            )
            .ok()
    }
}

fn create_not_found_response(
    environment: &ICoreWebView2Environment,
) -> Option<ICoreWebView2WebResourceResponse> {
    create_web_resource_response(
        environment,
        b"Not found",
        404,
        "Not Found",
        "text/plain; charset=utf-8",
        PASSIVE_FILE_HEADERS,
    )
}

fn create_virtual_resource_response(
    environment: &ICoreWebView2Environment,
    file_path: &Path,
    dark_mode: bool,
) -> Option<ICoreWebView2WebResourceResponse> {
    if links::is_markdown(file_path) {
        let (html_body, remote_images, bars) = if let Some(mark) = download_gate(file_path) {
            // Downloaded from the internet (#24): nothing from the document
            // is shown until the user chooses to open it.
            let gate = markdown::download_gate_html(&file_path.display().to_string(), mark.host_url.as_deref());
            (gate, markdown::RemoteImages::Blocked, vec![markdown::download_gate_bar()])
        } else {
            match files::read_text(file_path) {
                Ok(markdown_content) => {
                    let remote_images = remote_images_for(file_path);
                    let bars = DOCUMENT_CHOICES.with(|c| {
                        let mut choices = c.borrow_mut();
                        choices.page_served(file_path);
                        let ask_network_files = settings::network_files_setting() == settings::AskSetting::Ask
                            && !choices.network_loaded(file_path);
                        markdown::document_bars(&markdown::PageBars {
                            remote_images,
                            ask_to_register: ask_to_register(),
                            ask_network_files,
                            closed: choices.closed_bars(file_path),
                        })
                    });
                    (markdown::markdown_to_page_html(&markdown_content, file_path), remote_images, bars)
                }
                Err(reason) => (
                    markdown::read_error_html(&file_path.display().to_string(), &reason),
                    markdown::RemoteImages::Blocked,
                    Vec::new(),
                ),
            }
        };
        PAGE_ALLOWS_REMOTE_IMAGES.with(|a| a.set(remote_images == markdown::RemoteImages::Allowed));
        let full_html = markdown::wrap_html(&html_body, dark_mode, remote_images, &bars);
        PAGE_BARS.with(|b| *b.borrow_mut() = bars);
        create_web_resource_response(
            environment,
            full_html.as_bytes(),
            200,
            "OK",
            "text/html; charset=utf-8",
            "",
        )
    } else {
        let bytes = match files::read_bytes(file_path) {
            Ok(bytes) => bytes,
            Err(reason) => {
                return create_web_resource_response(
                    environment,
                    format!("Could not read {}: {reason}", file_path.display()).as_bytes(),
                    500,
                    "Could not read",
                    "text/plain; charset=utf-8",
                    PASSIVE_FILE_HEADERS,
                );
            }
        };
        create_web_resource_response(
            environment,
            &bytes,
            200,
            "OK",
            content_type_for_path(file_path),
            PASSIVE_FILE_HEADERS,
        )
    }
}

/// Answer for a request MDView does not allow to reach the network.
fn create_blocked_response(
    environment: &ICoreWebView2Environment,
) -> Option<ICoreWebView2WebResourceResponse> {
    create_web_resource_response(
        environment,
        b"",
        403,
        concat!("Blocked by ", app_name!()),
        "text/plain; charset=utf-8",
        PASSIVE_FILE_HEADERS,
    )
}

/// A request outside MDView's own files may only go out if it is an image on
/// a page whose remote images the user chose to load.
fn remote_request_allowed(uri: &str, context: COREWEBVIEW2_WEB_RESOURCE_CONTEXT) -> bool {
    let lower = uri.to_ascii_lowercase();
    context == COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE
        && (lower.starts_with("https://") || lower.starts_with("http://"))
        && PAGE_ALLOWS_REMOTE_IMAGES.with(|a| a.get())
}

/// The mark of a downloaded document that MDView must ask about before
/// showing it: not if the user chose "Don't ask again" or already chose Open
/// while this document is open.
fn download_gate(file_path: &Path) -> Option<zone::Mark> {
    if settings::open_downloaded_without_asking() {
        return None;
    }
    let mark = zone::mark_of_the_web(file_path)?;
    let opened = DOCUMENT_CHOICES.with(|c| c.borrow_mut().opened_download(file_path));
    (!opened).then_some(mark)
}

/// Remote images for a document: the saved choice, or "load once" if the user
/// chose that for this document in the banner.
fn remote_images_for(file_path: &Path) -> markdown::RemoteImages {
    match settings::remote_images_setting() {
        settings::AskSetting::Always => markdown::RemoteImages::Allowed,
        settings::AskSetting::Never => markdown::RemoteImages::Blocked,
        settings::AskSetting::Ask => {
            let once = REMOTE_IMAGES_ONCE.with(|o| o.borrow().as_deref() == Some(file_path));
            if once {
                markdown::RemoteImages::Allowed
            } else {
                markdown::RemoteImages::AskToLoad
            }
        }
    }
}

/// The file a URL from the page stands for while the current file is shown,
/// or None if it does not name a plain file (see `vpath`).
fn resolve_virtual_mdview_url(url: &str) -> Option<String> {
    CURRENT_FILE.with(|f| {
        let current = f.borrow();
        vpath::resolve(url, Path::new(current.as_ref()?)).map(|path| path.to_string_lossy().into_owned())
    })
}

fn create_response_for_virtual_url(
    environment: &ICoreWebView2Environment,
    url: &str,
) -> Option<ICoreWebView2WebResourceResponse> {
    // A URL that names no plain file is answered here, never sent anywhere.
    let Some(resolved_path) = resolve_virtual_mdview_url(url) else {
        return create_not_found_response(environment);
    };
    let file_path = Path::new(&resolved_path);
    // Decided before the file is touched: MDView does not connect to a
    // network location until the user agrees.
    if !network_file_allowed(file_path) {
        return create_blocked_response(environment);
    }
    if !file_path.exists() {
        return create_not_found_response(environment);
    }
    create_virtual_resource_response(environment, file_path, is_windows_dark_mode())
}

/// Whether a file the page asks for may be read. Files on another share or a
/// network drive wait for the user's answer in the network bar.
fn network_file_allowed(path: &Path) -> bool {
    let Some(current) = CURRENT_FILE.with(|f| f.borrow().clone()) else { return false };
    let location = match vpath::location(path, Path::new(&current), vpath::drive_kind) {
        Some(vpath::Location::Network(location)) => location,
        Some(vpath::Location::SameRoot | vpath::Location::Local) => return true,
        None => return false,
    };
    match settings::network_files_setting() {
        settings::AskSetting::Always => true,
        settings::AskSetting::Never => false,
        settings::AskSetting::Ask => DOCUMENT_CHOICES.with(|c| {
            let mut choices = c.borrow_mut();
            let Some(document) = choices.document().map(Path::to_path_buf) else { return false };
            if choices.network_loaded(&document) {
                return true;
            }
            if choices.note_network_blocked(&document, &location) {
                if let Some(hwnd) = MAIN_HWND.with(|h| *h.borrow()) {
                    // Told to the page from the window, not from inside this
                    // WebView2 callback.
                    let _ = unsafe { PostMessageW(Some(hwnd), WM_MDVIEW_NETWORK, WPARAM(0), LPARAM(0)) };
                }
            }
            false
        }),
    }
}

/// The network location (another share or a network drive) `path` is on,
/// seen from the current document; None for local files.
fn network_location(path: &Path) -> Option<String> {
    let current = CURRENT_FILE.with(|f| f.borrow().clone())?;
    match vpath::location(path, Path::new(&current), vpath::drive_kind)? {
        vpath::Location::Network(location) => Some(location),
        vpath::Location::SameRoot | vpath::Location::Local => None,
    }
}
/// Tell the page which network locations were blocked, for its network bar.
/// The page ignores locations it already lists.
fn post_network_blocked() {
    let locations = DOCUMENT_CHOICES.with(|c| {
        let mut choices = c.borrow_mut();
        match choices.document().map(Path::to_path_buf) {
            Some(document) => choices.network_blocked(&document),
            None => Vec::new(),
        }
    });
    CONTROLLER.with(|c| {
        if let Some(controller) = c.borrow().as_ref() {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                for location in locations {
                    let json: Vec<u16> = markdown::network_blocked_message(&location)
                        .encode_utf16()
                        .chain(std::iter::once(0))
                        .collect();
                    let _ = unsafe { webview.PostWebMessageAsJson(PCWSTR(json.as_ptr())) };
                }
            }
        }
    });
}

/// The user's answer in the network bar.
fn handle_network_files_choice(hwnd: HWND, choice: &str, document: &Path) {
    let save = |setting| match settings::set_network_files_setting(setting) {
        Ok(()) => true,
        Err(e) => {
            show_message(Some(hwnd), &format!("Could not save the network files setting:\n\n{e}"), MB_OK | MB_ICONERROR);
            false
        }
    };
    match choice {
        "once" => {
            DOCUMENT_CHOICES.with(|c| c.borrow_mut().load_network(document));
            reload_page();
        }
        "always" => {
            if save(settings::AskSetting::Always) {
                reload_page();
            }
        }
        "never" => {
            save(settings::AskSetting::Never);
        }
        _ => {}
    }
}

/// Any of MDView's own URLs: served by the host, never fetched from the web.
fn is_internal_virtual_url(url: &str) -> bool {
    crate::vpath::is_virtual_url(url)
}

/// Only MDView's own pages may be shown: Markdown documents on the virtual
/// origin, and the welcome page MDView itself loads with NavigateToString.
/// Everything else (script or <meta refresh> navigation, other files, other
/// sites) is cancelled; a link the user clicked goes through `queue_link`.
fn navigation_allowed(url: &str) -> bool {
    if crate::vpath::is_document_origin(url) {
        return resolve_virtual_mdview_url(url)
            .is_some_and(|path| links::is_markdown(Path::new(&path)));
    }
    if url.starts_with("data:") || url == "about:blank" {
        return EXPECT_STRING_PAGE.with(|e| e.replace(false));
    }
    false
}

/// Whether a message comes from a page MDView itself loaded.
fn is_trusted_page(source: &str) -> bool {
    if crate::vpath::is_document_origin(source) {
        return resolve_virtual_mdview_url(source)
            .is_some_and(|path| links::is_markdown(Path::new(&path)));
    }
    source.starts_with("data:") || source == "about:blank"
}

/// Handle a link outside of the WebView2 event that reported it, so the
/// confirmation dialog is not shown from inside a WebView2 callback.
fn queue_link(url: String) {
    PENDING_LINKS.with(|q| q.borrow_mut().push_back(url));
    if let Some(hwnd) = MAIN_HWND.with(|h| *h.borrow()) {
        let _ = unsafe { PostMessageW(Some(hwnd), WM_MDVIEW_LINK, WPARAM(0), LPARAM(0)) };
    }
}

fn handle_queued_links(hwnd: HWND) {
    while let Some(url) = PENDING_LINKS.with(|q| q.borrow_mut().pop_front()) {
        let action = links::classify_url(&url, |u| {
            resolve_virtual_mdview_url(u).map(std::path::PathBuf::from)
        }, network_location);
        links::perform(action, hwnd, |path| {
            load_file_into_webview(hwnd, &path.to_string_lossy())
        });
    }
}

/// A message from MDView's page script.
fn handle_page_message(hwnd: HWND, source: &str, msg: &std::collections::HashMap<String, String>) {
    match msg.get("type").map(String::as_str) {
        Some("close") => {
            let _ = unsafe { PostMessageW(Some(hwnd), WM_CLOSE, WPARAM(0), LPARAM(0)) };
        }
        Some("openLink") => {
            if let Some(url) = msg.get("url") {
                queue_link(url.clone());
            }
        }
        Some("networkStatus") => post_network_blocked(),
        Some("bar") => {
            let (Some(bar), Some(choice)) = (msg.get("bar"), msg.get("choice")) else { return };
            if !PAGE_BARS.with(|b| markdown::bar_choice_offered(&b.borrow(), bar, choice)) {
                return;
            }
            if choice == markdown::CLOSE_CHOICE {
                if let Some(path) = resolve_virtual_mdview_url(source) {
                    DOCUMENT_CHOICES.with(|c| c.borrow_mut().close_bar(Path::new(&path), bar));
                }
                return;
            }
            match bar.as_str() {
                markdown::DOWNLOAD_GATE => {
                    if let Some(path) = resolve_virtual_mdview_url(source) {
                        handle_download_gate_choice(hwnd, choice, Path::new(&path));
                    }
                }
                markdown::REMOTE_IMAGES_BAR => {
                    if let Some(path) = resolve_virtual_mdview_url(source) {
                        handle_remote_images_choice(hwnd, choice, Path::new(&path));
                    }
                }
                markdown::NETWORK_FILES_BAR => {
                    if let Some(path) = resolve_virtual_mdview_url(source) {
                        handle_network_files_choice(hwnd, choice, Path::new(&path));
                    }
                }
                markdown::REGISTER_BAR => handle_register_choice(hwnd, choice),
                _ => {}
            }
        }
        _ => {}
    }
}

/// The user's answer on the page shown instead of a downloaded document
/// (#24). The document is shown only if the choice worked.
fn handle_download_gate_choice(hwnd: HWND, choice: &str, document: &Path) {
    let result = match choice {
        "open" => {
            DOCUMENT_CHOICES.with(|c| c.borrow_mut().open_download(document));
            Ok(())
        }
        "unblock" => zone::remove_mark(document).map_err(|e| format!("Could not remove the mark of the web from\n{e}")),
        "never" => settings::set_open_downloaded_without_asking()
            .map_err(|e| format!("Could not save \"Don't ask again\":\n\n{e}")),
        "closeWindow" => {
            let _ = unsafe { PostMessageW(Some(hwnd), WM_CLOSE, WPARAM(0), LPARAM(0)) };
            return;
        }
        _ => return,
    };
    match result {
        Ok(()) => reload_page(),
        Err(message) => show_message(Some(hwnd), &message, MB_OK | MB_ICONERROR),
    }
}

fn reload_page() {
    CONTROLLER.with(|c| {
        if let Some(controller) = c.borrow().as_ref() {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                let _ = unsafe { webview.Reload() };
            }
        }
    });
}

/// The user's answer in the remote-images banner.
fn handle_remote_images_choice(hwnd: HWND, choice: &str, document: &Path) {
    let reload = match choice {
        "once" => {
            REMOTE_IMAGES_ONCE.with(|o| *o.borrow_mut() = Some(document.to_path_buf()));
            true
        }
        "always" => save_remote_images_setting(hwnd, settings::AskSetting::Always),
        "never" => {
            save_remote_images_setting(hwnd, settings::AskSetting::Never);
            false
        }
        _ => false,
    };
    if reload {
        reload_page();
    }
}

fn save_remote_images_setting(hwnd: HWND, setting: settings::AskSetting) -> bool {
    match settings::set_remote_images_setting(setting) {
        Ok(()) => true,
        Err(e) => {
            let text: Vec<u16> = format!("Could not save the remote images setting:\n\n{e}\0")
                .encode_utf16()
                .collect();
            let title: Vec<u16> = concat!(app_name!(), "\0").encode_utf16().collect();
            unsafe {
                MessageBoxW(Some(hwnd), PCWSTR(text.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
            }
            false
        }
    }
}

/// Load a file into the WebView
fn load_file_into_webview(hwnd: HWND, file_path: &str) {
    if fs::metadata(file_path).is_err() {
        let msg: Vec<u16> = format!("Failed to read file:\n{}\0", file_path)
            .encode_utf16()
            .collect();
        let title: Vec<u16> = "Error\0".encode_utf16().collect();
        unsafe {
            MessageBoxW(
                Some(hwnd),
                PCWSTR(msg.as_ptr()),
                PCWSTR(title.as_ptr()),
                MB_OK | MB_ICONERROR,
            );
        }
        return;
    }

    if let Some(url) = virtual_url_for_file(Path::new(file_path)) {
        CONTROLLER.with(|c| {
            if let Some(controller) = c.borrow().as_ref() {
                unsafe {
                    if let Ok(webview) = controller.CoreWebView2() {
                        let url_wide = pwstr_from_str(&url);
                        let _ = webview.Navigate(PCWSTR(url_wide.as_ptr()));
                    }
                }
            }
        });
    }

    // Update window title
    let filename = Path::new(file_path)
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| file_path.to_string());
    let title: Vec<u16> = format!("{filename} - {APP_NAME}\0").encode_utf16().collect();
    unsafe {
        let _ = SetWindowTextW(hwnd, PCWSTR(title.as_ptr()));
    }

    // Store current file and add to recent
    CURRENT_FILE.with(|f| {
        *f.borrow_mut() = Some(file_path.to_string());
    });
    add_to_recent_files(file_path);
}

/// Handle dropped files
fn handle_drop_files(hwnd: HWND, hdrop: HDROP) {
    unsafe {
        // Get the first dropped file
        let mut buffer = [0u16; 260];
        let len = DragQueryFileW(hdrop, 0, Some(&mut buffer));
        DragFinish(hdrop);

        if len > 0 {
            let path = String::from_utf16_lossy(&buffer[..len as usize]);
            // Check if it's a markdown file
            let ext = Path::new(&path)
                .extension()
                .map(|e| e.to_string_lossy().to_lowercase())
                .unwrap_or_default();
            if ext == "md" || ext == "markdown" {
                load_file_into_webview(hwnd, &path);
            } else {
                let msg: Vec<u16> = "Only Markdown files (.md, .markdown) are supported.\0"
                    .encode_utf16()
                    .collect();
                let title: Vec<u16> = "Unsupported File\0".encode_utf16().collect();
                MessageBoxW(
                    Some(hwnd),
                    PCWSTR(msg.as_ptr()),
                    PCWSTR(title.as_ptr()),
                    MB_OK | MB_ICONWARNING,
                );
            }
        }
    }
}

// ============================================================================
// File Association Registration
// ============================================================================

/// Helper: create a registry key under HKCU and set its default (unnamed) value
fn reg_create_default(subkey_path: &str, value: &str) -> Result<(), String> {
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("{}\0", subkey_path).encode_utf16().collect();
        let mut disposition = REG_CREATE_KEY_DISPOSITION::default();

        if RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            None,
            windows::Win32::System::Registry::REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut hkey,
            Some(&mut disposition),
        )
        .is_err()
        {
            return Err(format!("Failed to create key: {}", subkey_path));
        }

        let data: Vec<u16> = format!("{}\0", value).encode_utf16().collect();
        let result = RegSetValueExW(
            hkey,
            PCWSTR::null(),
            Some(0),
            REG_SZ,
            Some(std::slice::from_raw_parts(
                data.as_ptr() as *const u8,
                data.len() * 2,
            )),
        );
        let _ = RegCloseKey(hkey);

        if result.is_err() {
            return Err(format!("Failed to set value for: {}", subkey_path));
        }
        Ok(())
    }
}

/// Helper: create a registry key under HKCU and set a named string value
fn reg_set_named(subkey_path: &str, name: &str, value: &str) -> Result<(), String> {
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("{}\0", subkey_path).encode_utf16().collect();
        let mut disposition = REG_CREATE_KEY_DISPOSITION::default();

        if RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            None,
            windows::Win32::System::Registry::REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut hkey,
            Some(&mut disposition),
        )
        .is_err()
        {
            return Err(format!("Failed to create key: {}", subkey_path));
        }

        let name_wide: Vec<u16> = format!("{}\0", name).encode_utf16().collect();
        let data: Vec<u16> = format!("{}\0", value).encode_utf16().collect();
        let result = RegSetValueExW(
            hkey,
            PCWSTR(name_wide.as_ptr()),
            Some(0),
            REG_SZ,
            Some(std::slice::from_raw_parts(
                data.as_ptr() as *const u8,
                data.len() * 2,
            )),
        );
        let _ = RegCloseKey(hkey);

        if result.is_err() {
            return Err(format!(
                "Failed to set value '{}' for: {}",
                name, subkey_path
            ));
        }
        Ok(())
    }
}

/// Register MDView as a handler for .md/.markdown files in the current user's registry
fn register_file_association() -> Result<(), String> {
    let exe_path = std::env::current_exe()
        .map_err(|e| format!("Failed to get executable path: {}", e))?
        .to_string_lossy()
        .to_string();

    // ProgID: Software\Classes\MDView.Markdown
    reg_create_default(
        &format!("Software\\Classes\\{}", PROGID),
        PROGID_DESCRIPTION,
    )?;
    reg_create_default(
        &format!("Software\\Classes\\{}\\DefaultIcon", PROGID),
        &format!("\"{}\",0", exe_path),
    )?;
    reg_create_default(
        &format!("Software\\Classes\\{}\\shell\\open\\command", PROGID),
        &format!("\"{}\" \"%1\"", exe_path),
    )?;

    // OpenWithProgids for .md and .markdown
    for ext in &[".md", ".markdown"] {
        reg_set_named(
            &format!("Software\\Classes\\{}\\OpenWithProgids", ext),
            PROGID,
            "",
        )?;
    }

    // Capabilities
    reg_set_named(CAPABILITIES_KEY, "ApplicationDescription", APP_DESCRIPTION)?;
    reg_set_named(CAPABILITIES_KEY, "ApplicationName", APP_NAME)?;
    reg_set_named(FILE_ASSOCIATIONS_KEY, ".md", PROGID)?;
    reg_set_named(FILE_ASSOCIATIONS_KEY, ".markdown", PROGID)?;

    // RegisteredApplications
    reg_set_named(
        "Software\\RegisteredApplications",
        REGISTERED_APP_NAME,
        CAPABILITIES_KEY,
    )?;

    // App Paths
    reg_create_default(
        "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\mdview.exe",
        &exe_path,
    )?;

    // Notify shell of association change
    unsafe {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    Ok(())
}

/// Remove MDView file association from the current user's registry
fn unregister_file_association() -> Result<(), String> {
    unsafe {
        // Delete ProgID key tree
        let subkey: Vec<u16> = format!("Software\\Classes\\{}\0", PROGID)
            .encode_utf16()
            .collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(subkey.as_ptr()));

        // Remove from OpenWithProgids for .md and .markdown
        for ext in &[".md", ".markdown"] {
            let mut hkey = std::mem::zeroed();
            let owp_path: Vec<u16> = format!("Software\\Classes\\{}\\OpenWithProgids\0", ext)
                .encode_utf16()
                .collect();
            if RegOpenKeyExW(
                HKEY_CURRENT_USER,
                PCWSTR(owp_path.as_ptr()),
                Some(0),
                KEY_WRITE,
                &mut hkey,
            )
            .is_ok()
            {
                let name: Vec<u16> = format!("{}\0", PROGID).encode_utf16().collect();
                let _ = RegDeleteValueW(hkey, PCWSTR(name.as_ptr()));
                let _ = RegCloseKey(hkey);
            }
        }

        // Delete Capabilities tree
        let cap_key: Vec<u16> = concat!("Software\\", app_name!(), "\\Capabilities\0").encode_utf16().collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(cap_key.as_ptr()));

        // Remove from RegisteredApplications
        {
            let mut hkey = std::mem::zeroed();
            let ra_path: Vec<u16> = "Software\\RegisteredApplications\0"
                .encode_utf16()
                .collect();
            if RegOpenKeyExW(
                HKEY_CURRENT_USER,
                PCWSTR(ra_path.as_ptr()),
                Some(0),
                KEY_WRITE,
                &mut hkey,
            )
            .is_ok()
            {
                let name: Vec<u16> = format!("{REGISTERED_APP_NAME}\0").encode_utf16().collect();
                let _ = RegDeleteValueW(hkey, PCWSTR(name.as_ptr()));
                let _ = RegCloseKey(hkey);
            }
        }

        // Delete App Paths
        let ap_key: Vec<u16> =
            "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\mdview.exe\0"
                .encode_utf16()
                .collect();
        let _ = RegDeleteTreeW(HKEY_CURRENT_USER, PCWSTR(ap_key.as_ptr()));

        // Notify shell
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None);
    }

    Ok(())
}

/// Check if MDView is registered as a file handler and the command points to the current exe
fn is_file_association_registered() -> bool {
    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = format!("Software\\Classes\\{}\\shell\\open\\command\0", PROGID)
            .encode_utf16()
            .collect();

        if RegOpenKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            KEY_READ,
            &mut hkey,
        )
        .is_err()
        {
            return false;
        }

        let mut data = vec![0u16; 1024];
        let mut data_size = (data.len() * 2) as u32;
        let mut data_type = REG_VALUE_TYPE::default();

        let result = RegQueryValueExW(
            hkey,
            PCWSTR::null(),
            None,
            Some(&mut data_type),
            Some(data.as_mut_ptr() as *mut u8),
            Some(&mut data_size),
        );
        let _ = RegCloseKey(hkey);

        if result.is_err() {
            return false;
        }

        let len = data.iter().position(|&c| c == 0).unwrap_or(data.len());
        let value = String::from_utf16_lossy(&data[..len]);

        if let Ok(exe_path) = std::env::current_exe() {
            let exe_str = exe_path.to_string_lossy().to_lowercase();
            value.to_lowercase().contains(&exe_str)
        } else {
            false
        }
    }
}

const DONT_ASK_REGISTER: &str = "DontAskRegister";

/// Whether to ask about registering MDView as .md viewer (#37): only while it
/// is not registered, the user didn't choose "Don't ask again", and didn't
/// choose "Not now" in this session.
fn should_ask_to_register(registered: bool, dont_ask: bool, dismissed: bool) -> bool {
    !registered && !dont_ask && !dismissed
}

fn ask_to_register() -> bool {
    should_ask_to_register(
        is_file_association_registered(),
        settings::read_dword(DONT_ASK_REGISTER) == Some(1),
        REGISTER_BAR_DISMISSED.with(|d| d.get()),
    )
}

/// The user's answer in the registration bar.
fn handle_register_choice(hwnd: HWND, choice: &str) {
    match choice {
        "register" => register_and_open_settings(Some(hwnd)),
        "notNow" => REGISTER_BAR_DISMISSED.with(|d| d.set(true)),
        "never" => {
            if let Err(e) = settings::write_dword(DONT_ASK_REGISTER, 1) {
                show_message(
                    Some(hwnd),
                    &format!("Could not save \"Don't ask again\":\n\n{e}"),
                    MB_OK | MB_ICONERROR,
                );
            }
        }
        _ => {}
    }
}
/// The Settings page for choosing MDView for .md files. Windows 11 opens
/// MDView's own default-apps page; versions without that deep link open the
/// general Default apps page.
fn default_apps_settings_uri() -> String {
    format!("ms-settings:defaultapps?registeredAppUser={REGISTERED_APP_NAME}")
}

/// After registering, take the user to the Settings page where MDView can be
/// chosen for .md files.
///
/// Not IApplicationAssociationRegistrationUI::LaunchAdvancedAssociationUI:
/// Windows 10 and 11 only show "To change your default apps, go to Settings >
/// Apps > Default apps." for it, and report success (#36).
fn open_association_settings(owner: Option<HWND>) {
    unsafe {
        let settings_uri = U16CString::from_str(default_apps_settings_uri()).unwrap_or_default();
        let open_verb = U16CString::from_str("open").unwrap_or_default();
        let result = ShellExecuteW(
            owner,
            PCWSTR(open_verb.as_ptr()),
            PCWSTR(settings_uri.as_ptr()),
            None,
            None,
            SW_SHOWNORMAL,
        );
        // ShellExecute returns a value greater than 32 on success.
        if result.0 as usize <= 32 {
            show_message(
                owner,
                "MDView is registered, but Windows Settings could not be opened.\n\n\
                 To use MDView for .md files, open Settings > Apps > Default apps, \
                 search for MDView and choose it for .md and .markdown.",
                MB_OK | MB_ICONINFORMATION,
            );
        }
    }
}

/// Register MDView and open Settings; tell the user if registering failed.
fn register_and_open_settings(owner: Option<HWND>) {
    match register_file_association() {
        Ok(()) => open_association_settings(owner),
        Err(e) => show_message(
            owner,
            &format!("Could not register {APP_NAME} as a viewer for .md files:\n\n{e}"),
            MB_OK | MB_ICONERROR,
        ),
    }
}

fn show_message(owner: Option<HWND>, text: &str, style: MESSAGEBOX_STYLE) {
    let text: Vec<u16> = format!("{text}\0").encode_utf16().collect();
    let title: Vec<u16> = concat!(app_name!(), "\0").encode_utf16().collect();
    unsafe {
        MessageBoxW(owner, PCWSTR(text.as_ptr()), PCWSTR(title.as_ptr()), style);
    }
}

fn run_gui(title: &str, html: &str, file_path: Option<&str>) -> windows::core::Result<()> {
    unsafe {
        // Initialize COM for this thread (required for WebView2)
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(windows::core::Error::from(hr));
        }

        let result = run_gui_inner(title, html, file_path);

        CoUninitialize();
        result
    }
}

fn run_gui_inner(title: &str, html: &str, file_path: Option<&str>) -> windows::core::Result<()> {
    unsafe {
        // Enable dark mode for menus BEFORE creating any windows
        if is_windows_dark_mode() {
            set_preferred_app_mode_dark();
        }

        // Load recent files from registry
        let recent = load_recent_files();
        RECENT_FILES.with(|files| {
            *files.borrow_mut() = recent;
        });

        // Register window class
        let class_name_wide: Vec<u16> = WINDOW_CLASS
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();
        let hinstance = GetModuleHandleW(None)?;

        // Load application icon from resources
        let icon = LoadIconW(Some(hinstance.into()), PCWSTR(1 as *const u16)).ok();

        // Use dark background brush for dark mode
        let bg_brush = if is_windows_dark_mode() {
            CreateSolidBrush(COLORREF(DARK_CLIENT))
        } else {
            HBRUSH((COLOR_WINDOW.0 + 1) as _)
        };

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: bg_brush,
            lpszClassName: PCWSTR(class_name_wide.as_ptr()),
            hIcon: icon.unwrap_or_default(),
            hIconSm: icon.unwrap_or_default(),
            ..Default::default()
        };
        RegisterClassExW(&wc);

        // Create menu bar
        let menu = create_menu()?;

        // Create accelerator table
        let haccel = create_accelerators()?;
        ACCEL_HANDLE.with(|a| {
            *a.borrow_mut() = Some(haccel);
        });

        // Create main window with native menu bar (using SetMenu via CreateWindowExW)
        let title_wide: Vec<u16> = format!("{title} - {APP_NAME}")
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        let hwnd = CreateWindowExW(
            WS_EX_ACCEPTFILES, // Enable drag & drop
            PCWSTR(class_name_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            1024,
            768,
            None,
            Some(menu), // Native menu bar via SetMenu
            Some(hinstance.into()),
            None,
        )?;

        // Store menu handle for recent files updates
        MENU_HANDLE.with(|m| {
            *m.borrow_mut() = Some(menu);
        });

        // Store main window handle
        MAIN_HWND.with(|h| {
            *h.borrow_mut() = Some(hwnd);
        });

        // Store current file path
        if let Some(path) = file_path {
            CURRENT_FILE.with(|f| {
                *f.borrow_mut() = Some(path.to_string());
            });
            add_to_recent_files(path);
        }

        // Update recent files menu
        update_recent_files_menu();

        // Set dark mode title bar and menu if Windows is in dark mode
        if is_windows_dark_mode() {
            // Allow dark mode for this window (needed for dark popup menus)
            allow_dark_mode_for_window(hwnd);

            // Enable immersive dark mode for title bar
            let use_dark_mode: u32 = 1;
            let _ = DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                &use_dark_mode as *const u32 as *const std::ffi::c_void,
                std::mem::size_of::<u32>() as u32,
            );

            // Force menu bar redraw to trigger UAH messages
            let _ = DrawMenuBar(hwnd);
        }

        // Enable rounded corners on Windows 11
        let corner_preference = DWMWCP_ROUND.0;
        let _ = DwmSetWindowAttribute(
            hwnd,
            DWMWA_WINDOW_CORNER_PREFERENCE,
            &corner_preference as *const i32 as *const std::ffi::c_void,
            std::mem::size_of::<i32>() as u32,
        );

        // Initialize WebView2
        init_webview2_gui(hwnd, html)?;

        // Restore window position (if saved), otherwise center on screen
        let position_restored = restore_window_position(hwnd);

        if !position_restored {
            // Center window on primary monitor - use 70% of screen size
            let screen_width = GetSystemMetrics(SM_CXSCREEN);
            let screen_height = GetSystemMetrics(SM_CYSCREEN);
            let window_width = (screen_width * 70) / 100;
            let window_height = (screen_height * 75) / 100;
            let x = (screen_width - window_width) / 2;
            let y = (screen_height - window_height) / 2;
            let _ = SetWindowPos(hwnd, None, x, y, window_width, window_height, SWP_NOZORDER);
        }

        // Show window
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);
        // Activation inside ShowWindow focuses this window before WebView2 can
        // take focus; hand it to the document now that the window is shown.
        focus_webview();

        // Message loop with accelerator support
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            // Check if accelerator handles this message
            if TranslateAcceleratorW(hwnd, haccel, &msg) == 0 {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }

        // Cleanup accelerator table
        let _ = DestroyAcceleratorTable(haccel);

        Ok(())
    }
}

thread_local! {
    static CONTROLLER: RefCell<Option<ICoreWebView2Controller>> = const { RefCell::new(None) };
    static CURRENT_FILE: RefCell<Option<String>> = const { RefCell::new(None) };
    static RECENT_FILES: RefCell<Vec<String>> = const { RefCell::new(Vec::new()) };
    static MENU_HANDLE: RefCell<Option<HMENU>> = const { RefCell::new(None) };
    static ACCEL_HANDLE: RefCell<Option<HACCEL>> = const { RefCell::new(None) };
    static MAIN_HWND: RefCell<Option<HWND>> = const { RefCell::new(None) };
    static ENVIRONMENT: RefCell<Option<ICoreWebView2Environment>> = const { RefCell::new(None) };
    /// Links waiting to be handled outside the WebView2 event that reported them.
    static PENDING_LINKS: RefCell<std::collections::VecDeque<String>> = const { RefCell::new(std::collections::VecDeque::new()) };
    /// Set just before MDView calls NavigateToString, so only that navigation
    /// to a data:/about:blank page is allowed.
    static EXPECT_STRING_PAGE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    /// The document for which the user chose "Load images" in the banner.
    static REMOTE_IMAGES_ONCE: RefCell<Option<std::path::PathBuf>> = const { RefCell::new(None) };
    /// Whether the page being shown may load remote images (set when MDView serves it).
    static PAGE_ALLOWS_REMOTE_IMAGES: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    /// The bars MDView put on the page being shown; only answers to these are accepted.
    static PAGE_BARS: RefCell<Vec<markdown::Bar>> = const { RefCell::new(Vec::new()) };
    /// "Not now" in the registration bar: don't ask again in this session.
    static REGISTER_BAR_DISMISSED: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
    /// What the user chose while the document being shown is open (#24).
    static DOCUMENT_CHOICES: RefCell<markdown::DocumentChoices> = RefCell::new(markdown::DocumentChoices::default());
}

/// Posted to the main window to handle queued links (see `queue_link`).
const WM_MDVIEW_LINK: u32 = WM_APP + 1;
/// A file from a network location was blocked: tell the page.
const WM_MDVIEW_NETWORK: u32 = WM_APP + 2;

/// Edit > Find: open WebView2's built-in find bar.
fn open_find(hwnd: HWND) {
    focus_webview();
    if let Err(e) = start_find(hwnd) {
        show_find_error(hwnd, &e);
    }
}

fn show_find_error(hwnd: HWND, e: &windows::core::Error) {
    let msg: Vec<u16> = format!("Could not open the find bar: {}\0", e.message())
    .encode_utf16()
    .collect();
    let title: Vec<u16> = concat!(app_name!(), "\0").encode_utf16().collect();
    unsafe {
        MessageBoxW(
            Some(hwnd),
            PCWSTR(msg.as_ptr()),
            PCWSTR(title.as_ptr()),
            MB_OK | MB_ICONWARNING,
        );
    }
}

fn start_find(hwnd: HWND) -> windows::core::Result<()> {
    let environment = ENVIRONMENT
        .with(|e| e.borrow().clone())
        .ok_or_else(|| windows::core::Error::new(E_FAIL, "WebView2 is not initialised"))?;
    let webview = CONTROLLER
        .with(|c| c.borrow().clone())
        .ok_or_else(|| windows::core::Error::new(E_FAIL, "WebView2 is not initialised"))?;
    unsafe {
        let find = webview.CoreWebView2()?.cast::<ICoreWebView2_28>()?.Find()?;
        let options = environment
            .cast::<ICoreWebView2Environment15>()?
            .CreateFindOptions()?;
        options.SetSuppressDefaultFindDialog(false)?;
        // An empty term opens the find bar without searching yet. It must be a
        // real empty string: pwstr_from_str("") returns a null pointer.
        let empty_term: [u16; 1] = [0];
        options.SetFindTerm(PCWSTR(empty_term.as_ptr()))?;
        let handler = webview2_com::FindStartCompletedHandler::create(Box::new(move |result| {
            if let Err(e) = result {
                show_find_error(hwnd, &e);
            }
            Ok(())
        }));
        find.Start(&options, &handler)
    }
}

/// Give keyboard focus to the document, so Ctrl+F, scrolling keys and Esc
/// work without clicking in it first.
fn focus_webview() {
    CONTROLLER.with(|c| {
        if let Some(controller) = c.borrow().as_ref() {
            let _ = unsafe { controller.MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC) };
        }
    });
}

fn init_webview2_gui(hwnd: HWND, html: &str) -> windows::core::Result<()> {
    let html_owned = html.to_string();
    let completed = Rc::new(RefCell::new(false));
    let error_result: Rc<RefCell<Option<windows::core::Error>>> = Rc::new(RefCell::new(None));

    let completed_clone = completed.clone();
    let error_clone = error_result.clone();

    let env_handler = CreateCoreWebView2EnvironmentCompletedHandler::create(Box::new(
        move |_error_code, environment| {
            let environment = match environment {
                Some(env) => env,
                None => {
                    *error_clone.borrow_mut() = Some(windows::core::Error::from(E_FAIL));
                    *completed_clone.borrow_mut() = true;
                    return Ok(());
                }
            };

            let html_for_nav = html_owned.clone();
            let completed_inner = completed_clone.clone();
            let error_inner = error_clone.clone();
            let environment_for_resources = environment.clone();
            ENVIRONMENT.with(|e| *e.borrow_mut() = Some(environment.clone()));

            let ctrl_handler = CreateCoreWebView2ControllerCompletedHandler::create(Box::new(
                move |_error_code, controller| {
                    let controller = match controller {
                        Some(ctrl) => ctrl,
                        None => {
                            *error_inner.borrow_mut() = Some(windows::core::Error::from(E_FAIL));
                            *completed_inner.borrow_mut() = true;
                            return Ok(());
                        }
                    };

                    unsafe {
                        // Make controller visible
                        let _ = controller.SetIsVisible(true);

                        // Set bounds to fill parent window
                        let mut rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut rect);
                        let _ = controller.SetBounds(rect);

                        // Disable external drop so our WM_DROPFILES handler works
                        if let Ok(controller4) = controller.cast::<ICoreWebView2Controller4>() {
                            let _ = controller4.SetAllowExternalDrop(false);
                        }

                        if let Ok(webview) = controller.CoreWebView2() {
                            if let Ok(settings) = webview.Settings() {
                                let _ = settings.SetIsScriptEnabled(true);
                                let _ = settings.SetIsWebMessageEnabled(true);
                                let _ = settings.SetAreDefaultContextMenusEnabled(true);
                                let _ = settings.SetIsStatusBarEnabled(false);
                                let _ = settings.SetAreHostObjectsAllowed(false);
                            }

                            // Every request goes through MDView: its own files are
                            // served from disk, and nothing else reaches the network
                            // except remote images the user chose to load. (A
                            // cancelled navigation, e.g. <meta refresh>, can already
                            // have sent its request; this answers it locally.)
                            let filter = pwstr_from_str("*");
                            let _ = webview.AddWebResourceRequestedFilter(
                                PCWSTR(filter.as_ptr()),
                                COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL,
                            );

                            let resource_environment = environment_for_resources.clone();
                            let resource_handler =
                                webview2_com::WebResourceRequestedEventHandler::create(Box::new(
                                    move |_webview, args| {
                                        if let Some(args) = args {
                                            if let Ok(request) = args.Request() {
                                                let mut uri_ptr: windows::core::PWSTR =
                                                    windows::core::PWSTR::null();
                                                if request.Uri(&mut uri_ptr).is_ok()
                                                    && !uri_ptr.is_null()
                                                {
                                                    let uri =
                                                        uri_ptr.to_string().unwrap_or_default();
                                                    windows::Win32::System::Com::CoTaskMemFree(
                                                        Some(uri_ptr.0 as *const _),
                                                    );
                                                    let response = if is_internal_virtual_url(&uri) {
                                                        create_response_for_virtual_url(
                                                            &resource_environment,
                                                            &uri,
                                                        )
                                                    } else {
                                                        let mut context =
                                                            COREWEBVIEW2_WEB_RESOURCE_CONTEXT::default();
                                                        let _ = args.ResourceContext(&mut context);
                                                        if remote_request_allowed(&uri, context) {
                                                            None
                                                        } else {
                                                            create_blocked_response(&resource_environment)
                                                        }
                                                    };
                                                    if let Some(response) = response {
                                                        let _ = args.SetResponse(&response);
                                                    }
                                                }
                                            }
                                        }
                                        Ok(())
                                    },
                                ));
                            let mut resource_token: i64 = 0;
                            let _ = webview
                                .add_WebResourceRequested(&resource_handler, &mut resource_token);

                            // Only MDView's own pages may load; a link the user
                            // clicked is handled by the link policy instead.
                            let nav_handler = webview2_com::NavigationStartingEventHandler::create(
                                Box::new(move |_webview, args| {
                                    if let Some(args) = args {
                                        let mut uri_ptr: windows::core::PWSTR =
                                            windows::core::PWSTR::null();
                                        let mut user_initiated = windows::core::BOOL(0);
                                        if args.Uri(&mut uri_ptr).is_ok() && !uri_ptr.is_null() {
                                            let uri = uri_ptr.to_string().unwrap_or_default();
                                            let _ = args.IsUserInitiated(&mut user_initiated);
                                            windows::Win32::System::Com::CoTaskMemFree(Some(
                                                uri_ptr.0 as *const _,
                                            ));
                                            if !navigation_allowed(&uri) {
                                                let _ = args.SetCancel(true);
                                                if user_initiated.as_bool() {
                                                    queue_link(uri);
                                                }
                                            }
                                        } else {
                                            let _ = args.SetCancel(true);
                                        }
                                    }
                                    Ok(())
                                }),
                            );
                            let mut nav_token: i64 = 0;
                            let _ = webview.add_NavigationStarting(&nav_handler, &mut nav_token);

                            // No new windows: a link the user opened in a new
                            // window goes through the link policy.
                            let new_window_handler =
                                webview2_com::NewWindowRequestedEventHandler::create(Box::new(
                                    move |_webview, args| {
                                        if let Some(args) = args {
                                            let _ = args.SetHandled(true);
                                            let mut user_initiated = windows::core::BOOL(0);
                                            let _ = args.IsUserInitiated(&mut user_initiated);
                                            let mut uri_ptr = windows::core::PWSTR::null();
                                            if args.Uri(&mut uri_ptr).is_ok() && !uri_ptr.is_null() {
                                                let uri = uri_ptr.to_string().unwrap_or_default();
                                                windows::Win32::System::Com::CoTaskMemFree(Some(
                                                    uri_ptr.0 as *const _,
                                                ));
                                                if user_initiated.as_bool() {
                                                    queue_link(uri);
                                                }
                                            }
                                        }
                                        Ok(())
                                    },
                                ));
                            let mut new_window_token: i64 = 0;
                            let _ = webview
                                .add_NewWindowRequested(&new_window_handler, &mut new_window_token);

                            // No downloads and no permissions (camera, location, ...).
                            if let Ok(webview4) = webview.cast::<ICoreWebView2_4>() {
                                let download_handler =
                                    webview2_com::DownloadStartingEventHandler::create(Box::new(
                                        |_webview, args| {
                                            if let Some(args) = args {
                                                let _ = args.SetCancel(true);
                                                let _ = args.SetHandled(true);
                                            }
                                            Ok(())
                                        },
                                    ));
                                let mut download_token: i64 = 0;
                                let _ = webview4
                                    .add_DownloadStarting(&download_handler, &mut download_token);
                            }
                            let permission_handler =
                                webview2_com::PermissionRequestedEventHandler::create(Box::new(
                                    |_webview, args| {
                                        if let Some(args) = args {
                                            let _ = args.SetState(COREWEBVIEW2_PERMISSION_STATE_DENY);
                                        }
                                        Ok(())
                                    },
                                ));
                            let mut permission_token: i64 = 0;
                            let _ = webview
                                .add_PermissionRequested(&permission_handler, &mut permission_token);

                            if let Some(current_file) = CURRENT_FILE.with(|f| f.borrow().clone()) {
                                if let Some(url) = virtual_url_for_file(Path::new(&current_file)) {
                                    let url_wide = pwstr_from_str(&url);
                                    let _ = webview.Navigate(PCWSTR(url_wide.as_ptr()));
                                }
                            } else {
                                let html_wide = pwstr_from_str(&html_for_nav);
                                EXPECT_STRING_PAGE.with(|e| e.set(true));
                                if let Err(e) = webview.NavigateToString(PCWSTR(html_wide.as_ptr()))
                                {
                                    eprintln!("NavigateToString failed: {:?}", e);
                                }
                            }

                            // Messages from MDView's own page script only.
                            let main_hwnd = hwnd;
                            let handler = webview2_com::WebMessageReceivedEventHandler::create(
                                Box::new(move |_webview, args| {
                                    if let Some(args) = args {
                                        let mut source_ptr = windows::core::PWSTR::null();
                                        let source = if args.Source(&mut source_ptr).is_ok()
                                            && !source_ptr.is_null()
                                        {
                                            let s = source_ptr.to_string().unwrap_or_default();
                                            windows::Win32::System::Com::CoTaskMemFree(Some(
                                                source_ptr.0 as *const _,
                                            ));
                                            s
                                        } else {
                                            String::new()
                                        };
                                        if !is_trusted_page(&source) {
                                            return Ok(());
                                        }
                                        let mut message_ptr: windows::core::PWSTR =
                                            windows::core::PWSTR::null();
                                        if args.WebMessageAsJson(&mut message_ptr).is_ok()
                                            && !message_ptr.is_null()
                                        {
                                            let json = message_ptr.to_string().unwrap_or_default();
                                            windows::Win32::System::Com::CoTaskMemFree(Some(
                                                message_ptr.0 as *const _,
                                            ));
                                            if let Some(msg) = links::parse_page_message(&json) {
                                                handle_page_message(main_hwnd, &source, &msg);
                                            }
                                        }
                                    }
                                    Ok(())
                                }),
                            );
                            let mut token: i64 = 0;
                            let _ = webview.add_WebMessageReceived(&handler, &mut token);
                        }

                        CONTROLLER.with(|c| {
                            *c.borrow_mut() = Some(controller);
                        });
                    }

                    *completed_inner.borrow_mut() = true;
                    Ok(())
                },
            ));

            unsafe {
                let _ = environment.CreateCoreWebView2Controller(hwnd, &ctrl_handler);
            }
            Ok(())
        },
    ));

    // Own WebView2 data folder for this host program (#33)
    let user_data_folder = match settings::webview2_data_folder() {
        Ok(folder) => folder.to_string_lossy().into_owned(),
        Err(e) => {
            let text: Vec<u16> = format!("{APP_NAME} cannot start its viewer:\n\n{e}\0").encode_utf16().collect();
            let title: Vec<u16> = concat!(app_name!(), "\0").encode_utf16().collect();
            unsafe {
                MessageBoxW(Some(hwnd), PCWSTR(text.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
            }
            return Err(windows::core::Error::new(E_FAIL, e));
        }
    };
    let user_data_wide = pwstr_from_str(&user_data_folder);
    unsafe {
        let _ = CreateCoreWebView2EnvironmentWithOptions(
            None,
            PCWSTR(user_data_wide.as_ptr()),
            None,
            &env_handler,
        );
    }

    // Pump messages until WebView2 is ready
    unsafe {
        while !*completed.borrow() {
            let mut msg = MSG::default();
            if GetMessageW(&mut msg, None, 0, 0).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            } else {
                break;
            }
        }
    }

    if let Some(err) = error_result.borrow_mut().take() {
        return Err(err);
    }

    Ok(())
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        // UAH messages for dark menu bar
        dark_menu::WM_UAHDRAWMENU => {
            return dark_menu::handle_uah_draw_menu(hwnd, lparam);
        }
        dark_menu::WM_UAHDRAWMENUITEM => {
            return dark_menu::handle_uah_draw_menu_item(hwnd, lparam);
        }
        WM_NCPAINT | WM_NCACTIVATE => {
            // Let Windows draw first, then paint over the menu bar separator line
            let result = unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            dark_menu::paint_menu_separator(hwnd);
            return result;
        }
        WM_ACTIVATE | WM_SIZE => {
            let result = unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            dark_menu::paint_menu_separator(hwnd);

            // Also resize WebView on WM_SIZE
            if msg == WM_SIZE {
                CONTROLLER.with(|c| {
                    if let Some(controller) = c.borrow().as_ref() {
                        let mut rect = RECT::default();
                        let _ = unsafe { GetClientRect(hwnd, &mut rect) };
                        let _ = unsafe { controller.SetBounds(rect) };
                    }
                });
            }
            return result;
        }
        WM_MDVIEW_LINK => {
            handle_queued_links(hwnd);
            LRESULT(0)
        }
        WM_MDVIEW_NETWORK => {
            post_network_blocked();
            LRESULT(0)
        }
        // Windows gives focus to this window when it is activated; pass it on
        // to the document.
        WM_SETFOCUS => {
            focus_webview();
            LRESULT(0)
        }
        WM_INITMENUPOPUP => {
            // Update register/unregister enabled state when menu opens
            let menu = HMENU(wparam.0 as _);
            let registered = is_file_association_registered();
            unsafe {
                let _ = EnableMenuItem(
                    menu,
                    IDM_FILE_REGISTER,
                    if registered {
                        MF_BYCOMMAND | MF_GRAYED
                    } else {
                        MF_BYCOMMAND
                    },
                );
                let _ = EnableMenuItem(
                    menu,
                    IDM_FILE_UNREGISTER,
                    if registered {
                        MF_BYCOMMAND
                    } else {
                        MF_BYCOMMAND | MF_GRAYED
                    },
                );
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xFFFF) as u32;
            match cmd_id {
                IDM_FILE_OPEN => {
                    if let Some(path) = show_open_file_dialog(hwnd) {
                        load_file_into_webview(hwnd, &path);
                    }
                    LRESULT(0)
                }
                IDM_FILE_EXIT => {
                    unsafe { PostMessageW(Some(hwnd), WM_CLOSE, WPARAM(0), LPARAM(0)).ok() };
                    LRESULT(0)
                }
                IDM_FILE_REGISTER => {
                    register_and_open_settings(Some(hwnd));
                    LRESULT(0)
                }
                IDM_FILE_UNREGISTER => {
                    let _ = unregister_file_association();
                    LRESULT(0)
                }
                IDM_HELP_ABOUT => {
                    show_about_dialog(hwnd);
                    LRESULT(0)
                }
                IDM_EDIT_FIND => {
                    open_find(hwnd);
                    LRESULT(0)
                }
                id if id >= IDM_FILE_RECENT_BASE && id < IDM_FILE_RECENT_BASE + 10 => {
                    let index = (id - IDM_FILE_RECENT_BASE) as usize;
                    RECENT_FILES.with(|files| {
                        let files = files.borrow();
                        if let Some(path) = files.get(index) {
                            let path = path.clone();
                            drop(files);
                            load_file_into_webview(hwnd, &path);
                        }
                    });
                    LRESULT(0)
                }
                _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
            }
        }
        WM_DROPFILES => {
            let hdrop = HDROP(wparam.0 as _);
            handle_drop_files(hwnd, hdrop);
            LRESULT(0)
        }
        WM_DESTROY => {
            // Save window position before closing
            save_window_position(hwnd);

            CONTROLLER.with(|c| {
                if let Some(controller) = c.borrow_mut().take() {
                    let _ = unsafe { controller.Close() };
                }
            });
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn showing(file: &str) {
        CURRENT_FILE.with(|f| *f.borrow_mut() = Some(file.to_string()));
    }

    // Compared as text: the path is shown to the user (link questions, the
    // window title, recent files), so it must use `\` only (#49).
    fn assert_resolves(url: &str, expected: &str) {
        let resolved = resolve_virtual_mdview_url(url)
            .unwrap_or_else(|| panic!("{url} did not resolve to a file"));
        assert_eq!(resolved, expected, "resolved {url}");
    }

    #[test]
    fn test_document_url_maps_back_with_backslashes_only() {
        showing(r"C:\docs\note.md");
        assert_eq!(
            virtual_url_for_file(Path::new(r"C:\docs\note.md")).as_deref(),
            Some("https://mdview.example/docs/note.md")
        );
        assert_resolves("https://mdview.example/docs/img/a.png?v=2#x", r"C:\docs\img\a.png");
        assert_resolves(
            "https://mdview.example/Dokumenty/pozn%C3%A1mky/sn%C3%ADmek%20obrazovky.png",
            r"C:\Dokumenty\poznámky\snímek obrazovky.png",
        );
        showing(r"\\server\share\docs\note.md");
        assert_resolves("https://mdview.example/docs/img/a.png", r"\\server\share\docs\img\a.png");
    }

    // A URL from a document may only name a file below the root it maps to.
    // Encoded separators, drive letters, `..`, devices, pipes and alternate
    // data streams must never become a path elsewhere.
    #[test]
    fn test_urls_cannot_escape_to_other_paths() {
        showing(r"C:\docs\note.md");
        for url in [
            "https://mdview.example/%5C%5Cserver%5Cshare%5Cx.png",
            "https://mdview.example/%2F%2Fserver%2Fshare%2Fx.png",
            "https://mdview.example/Q%3A%5Cx.png",
            "https://mdview.example/Q%3A/x.png",
            "https://mdview.example/docs/..%5C..%5Cx.png",
            "https://mdview.example/docs/%2E%2E/x.png",
            "https://mdview.example/%5C%5C.%5Cpipe%5Cp",
            "https://mdview.example/%5C%5C%3F%5CC%3A%5Cx.png",
            "https://mdview.example/docs/a.png:secret",
            "https://mdview.example/docs/a.png%3Asecret",
            "https://mdview.example/docs/CON",
            "https://mdview.example/docs/nul.txt",
            "https://mdview.example/docs/COM1.md",
            "https://mdview.example/docs/x.png.",
            "https://mdview.example/docs/x.png%20",
            "https://mdview.example/docs//x.png",
        ] {
            assert_eq!(resolve_virtual_mdview_url(url), None, "{url} must not resolve to a file");
        }
        // Decoded once only: a double-encoded `\` is a literal name on the
        // document's own drive, never a separator.
        assert_resolves("https://mdview.example/%255C%255Cserver%255Cshare%255Cx.png", r"C:\%5C%5Cserver%5Cshare%5Cx.png");
    }
    fn placement(show_cmd: SHOW_WINDOW_CMD, flags: WINDOWPLACEMENT_FLAGS) -> WINDOWPLACEMENT {
        WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            flags,
            showCmd: show_cmd.0 as u32,
            ..Default::default()
        }
    }

    // Issue #6: closing a minimised window must not make the next start minimised.
    #[test]
    fn test_closing_minimised_saves_normal() {
        let p = placement(SW_SHOWMINIMIZED, WINDOWPLACEMENT_FLAGS(0));
        assert_eq!(show_cmd_to_save(&p), SW_SHOWNORMAL.0 as u32);
    }

    #[test]
    fn test_closing_minimised_from_maximised_saves_maximised() {
        let p = placement(SW_SHOWMINIMIZED, WPF_RESTORETOMAXIMIZED);
        assert_eq!(show_cmd_to_save(&p), SW_SHOWMAXIMIZED.0 as u32);
    }

    #[test]
    fn test_closing_maximised_or_normal_saves_as_is() {
        let p = placement(SW_SHOWMAXIMIZED, WINDOWPLACEMENT_FLAGS(0));
        assert_eq!(show_cmd_to_save(&p), SW_SHOWMAXIMIZED.0 as u32);
        let p = placement(SW_SHOWNORMAL, WINDOWPLACEMENT_FLAGS(0));
        assert_eq!(show_cmd_to_save(&p), SW_SHOWNORMAL.0 as u32);
    }

    // #36: the Settings deep link must name exactly the application MDView
    // registers under HKCU\Software\RegisteredApplications.
    #[test]
    fn test_settings_link_uses_the_registered_name() {
        assert_eq!(default_apps_settings_uri(), "ms-settings:defaultapps?registeredAppUser=MDView");
        assert!(default_apps_settings_uri().ends_with(&format!("={REGISTERED_APP_NAME}")));
    }

    // #37: ask only while not registered, not opted out, not dismissed now.
    #[test]
    fn test_registration_is_asked_only_when_it_should() {
        assert!(should_ask_to_register(false, false, false));
        for (registered, dont_ask, dismissed) in [
            (true, false, false),
            (false, true, false),
            (false, false, true),
            (true, true, true),
        ] {
            assert!(
                !should_ask_to_register(registered, dont_ask, dismissed),
                "registered={registered} dont_ask={dont_ask} dismissed={dismissed}"
            );
        }
    }

    // The product name is defined once; everything else is derived from it.
    #[test]
    fn test_names_are_derived_from_the_product_name() {
        assert_eq!(APP_NAME, "MDView");
        assert_eq!(REGISTRY_KEY, "Software\\MDView");
        assert_eq!(PROGID, "MDView.Markdown");
        assert_eq!(REGISTERED_APP_NAME, APP_NAME);
        assert_eq!(CAPABILITIES_KEY, "Software\\MDView\\Capabilities");
        assert_eq!(FILE_ASSOCIATIONS_KEY, "Software\\MDView\\Capabilities\\FileAssociations");
        assert!(links::REFUSED_PROGRAM.starts_with("MDView does not open"));
    }

    #[test]
    fn test_restore_never_starts_minimised() {
        assert_eq!(show_cmd_to_restore(SW_SHOWMAXIMIZED.0), SW_SHOWMAXIMIZED.0 as u32);
        assert_eq!(show_cmd_to_restore(SW_SHOWNORMAL.0), SW_SHOWNORMAL.0 as u32);
        // Saved by older versions when closed while minimised.
        assert_eq!(show_cmd_to_restore(SW_SHOWMINIMIZED.0), SW_SHOWNORMAL.0 as u32);
        assert_eq!(show_cmd_to_restore(SW_MINIMIZE.0), SW_SHOWNORMAL.0 as u32);
    }
}
