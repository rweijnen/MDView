#![windows_subsystem = "windows"]

mod dark_menu;
mod markdown;
mod terminal;

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

    // Validate mutually exclusive options
    let cli_format_count = opts.html_full as u8 + opts.html_body as u8 + opts.plain_text as u8 + opts.terminal_mode as u8;
    if cli_format_count > 1 {
        return Err("Options --term, --html, --body, and --text are mutually exclusive".to_string());
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
        RegOpenKeyExW, RegQueryValueExW, RegCloseKey, HKEY_CURRENT_USER, KEY_READ, REG_VALUE_TYPE,
    };
    use windows::core::PCWSTR;

    unsafe {
        let mut hkey = std::mem::zeroed();
        let subkey: Vec<u16> = "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\0"
            .encode_utf16().collect();

        if RegOpenKeyExW(HKEY_CURRENT_USER, PCWSTR(subkey.as_ptr()), Some(0), KEY_READ, &mut hkey).is_ok() {
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
    use windows::Win32::System::Console::{AttachConsole, ATTACH_PARENT_PROCESS};

    unsafe { AttachConsole(ATTACH_PARENT_PROCESS).is_ok() }
}

/// Write text directly to console and print newline to "complete" the prompt line
fn write_console(text: &str) {
    use windows::Win32::System::Console::{GetStdHandle, WriteConsoleW, STD_OUTPUT_HANDLE};

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
        GetConsoleMode, SetConsoleMode, GetStdHandle,
        CONSOLE_MODE, ENABLE_VIRTUAL_TERMINAL_PROCESSING, STD_OUTPUT_HANDLE,
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
        SendInput, INPUT, INPUT_KEYBOARD, KEYBDINPUT, KEYEVENTF_KEYUP, VK_RETURN,
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
            fs::read_to_string(path)
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

    // If no file and running in CLI/terminal mode, show help and exit
    if opts.file_path.is_none() && !opts.gui_mode {
        print_usage_console();
        send_enter_key(); // Release command prompt
        std::process::exit(0);
    }

    if opts.gui_mode {
        // GUI mode - open window with WebView2
        let (title, full_html) = if let Some(ref path) = opts.file_path {
            let markdown_content = match read_input(Some(path)) {
                Ok(content) => content,
                Err(e) => {
                    eprintln!("Error reading input: {}", e);
                    std::process::exit(1);
                }
            };
            let html_body = markdown::markdown_to_html(&markdown_content);
            let dark_mode = is_windows_dark_mode();
            let full_html = markdown::wrap_html(&html_body, dark_mode);
            let title = Path::new(path)
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .to_string();
            (title, full_html)
        } else {
            // No file - show welcome screen
            let dark_mode = is_windows_dark_mode();
            let welcome_html = markdown::wrap_html(
                "<div style=\"text-align: center; margin-top: 100px; color: #888;\">\
                 <h1>MDView</h1>\
                 <p>Open a Markdown file using <strong>File &gt; Open</strong> (Ctrl+O)</p>\
                 <p>or drag and drop a .md file onto this window.</p>\
                 </div>",
                dark_mode,
            );
            ("MDView".to_string(), welcome_html)
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
            markdown::wrap_html(&html_body, dark_mode)
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
    pwstr_from_str, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler,
};
use windows::core::Interface;
use widestring::U16CString;
use windows::core::PCWSTR;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Dwm::{
    DwmSetWindowAttribute, DWMNCRP_ENABLED, DWMSBT_MAINWINDOW, DWMWA_NCRENDERING_POLICY,
    DWMWA_SYSTEMBACKDROP_TYPE, DWMWA_USE_IMMERSIVE_DARK_MODE, DWMWA_WINDOW_CORNER_PREFERENCE,
    DWMWCP_ROUND,
};
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Registry::{
    RegCloseKey, RegCreateKeyExW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW,
    HKEY_CURRENT_USER, KEY_READ, KEY_WRITE, REG_CREATE_KEY_DISPOSITION, REG_DWORD, REG_SZ,
    REG_VALUE_TYPE,
};
use windows::Win32::UI::Shell::{
    DragFinish, DragQueryFileW, FileOpenDialog, IFileOpenDialog, ShellExecuteW,
    FOS_FILEMUSTEXIST, FOS_PATHMUSTEXIST, HDROP, SIGDN_FILESYSPATH,
};
use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows::Win32::UI::WindowsAndMessaging::*;

const WINDOW_CLASS: &str = "MDViewWindow";

// Menu item IDs
const IDM_FILE_OPEN: u32 = 1001;
const IDM_FILE_EXIT: u32 = 1002;
const IDM_HELP_ABOUT: u32 = 2001;
const IDM_FILE_RECENT_BASE: u32 = 1100; // 1100-1109 for recent files

// Registry key for settings
const REGISTRY_KEY: &str = "Software\\MDView";

/// Enable dark mode for menus using undocumented uxtheme.dll API
/// This calls SetPreferredAppMode (ordinal 135) with AllowDark (1)
fn set_preferred_app_mode_dark() {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

    unsafe {
        let uxtheme: Vec<u16> = "uxtheme.dll\0".encode_utf16().collect();
        if let Ok(hmodule) = LoadLibraryW(PCWSTR(uxtheme.as_ptr())) {
            // Ordinal 135 = SetPreferredAppMode
            // Ordinal 136 = FlushMenuThemes
            if let Some(set_preferred_app_mode) =
                GetProcAddress(hmodule, windows::core::PCSTR(135_usize as *const u8))
            {
                // PreferredAppMode: 0=Default, 1=AllowDark, 2=ForceDark, 3=ForceLight
                let func: extern "system" fn(i32) -> i32 = std::mem::transmute(set_preferred_app_mode);
                func(1); // AllowDark
            }
            // Flush menu themes to apply the change
            if let Some(flush_menu_themes) =
                GetProcAddress(hmodule, windows::core::PCSTR(136_usize as *const u8))
            {
                let func: extern "system" fn() = std::mem::transmute(flush_menu_themes);
                func();
            }
        }
    }
}

/// Create the application menu bar
fn create_menu() -> windows::core::Result<HMENU> {
    unsafe {
        let menu_bar = CreateMenu()?;
        let file_menu = CreatePopupMenu()?;
        let help_menu = CreatePopupMenu()?;

        // File menu items
        let open_text: Vec<u16> = "&Open\tCtrl+O\0".encode_utf16().collect();
        AppendMenuW(file_menu, MF_STRING, IDM_FILE_OPEN as usize, PCWSTR(open_text.as_ptr()))?;

        // Separator before recent files
        AppendMenuW(file_menu, MF_SEPARATOR, 0, None)?;

        // Recent files placeholder - will be populated later
        let recent_text: Vec<u16> = "Recent Files\0".encode_utf16().collect();
        AppendMenuW(file_menu, MF_STRING | MF_GRAYED, 0, PCWSTR(recent_text.as_ptr()))?;

        // Separator before exit
        AppendMenuW(file_menu, MF_SEPARATOR, 0, None)?;

        let exit_text: Vec<u16> = "E&xit\0".encode_utf16().collect();
        AppendMenuW(file_menu, MF_STRING, IDM_FILE_EXIT as usize, PCWSTR(exit_text.as_ptr()))?;

        // Help menu items
        let about_text: Vec<u16> = "&About MDView\0".encode_utf16().collect();
        AppendMenuW(help_menu, MF_STRING, IDM_HELP_ABOUT as usize, PCWSTR(about_text.as_ptr()))?;

        // Add submenus to menu bar
        let file_text: Vec<u16> = "&File\0".encode_utf16().collect();
        AppendMenuW(menu_bar, MF_POPUP, file_menu.0 as usize, PCWSTR(file_text.as_ptr()))?;

        let help_text: Vec<u16> = "&Help\0".encode_utf16().collect();
        AppendMenuW(menu_bar, MF_POPUP, help_menu.0 as usize, PCWSTR(help_text.as_ptr()))?;

        Ok(menu_bar)
    }
}

/// Create accelerator table for keyboard shortcuts
fn create_accelerators() -> windows::core::Result<HACCEL> {
    unsafe {
        let accels = [
            ACCEL {
                fVirt: FVIRTKEY | FCONTROL,
                key: 'O' as u16,
                cmd: IDM_FILE_OPEN as u16,
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
        let title: Vec<u16> = "About MDView\0".encode_utf16().collect();
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
                        // Show disabled placeholder
                        let text: Vec<u16> = "(No Recent Files)\0".encode_utf16().collect();
                        let _ = InsertMenuW(
                            file_menu,
                            2,
                            MF_BYPOSITION | MF_STRING | MF_GRAYED,
                            0,
                            PCWSTR(text.as_ptr()),
                        );
                    } else {
                        // Add recent files
                        for (i, path) in files.iter().enumerate() {
                            let filename = Path::new(path)
                                .file_name()
                                .map(|n| n.to_string_lossy().to_string())
                                .unwrap_or_else(|| path.clone());
                            let text: Vec<u16> =
                                format!("&{} {}\0", i + 1, filename).encode_utf16().collect();
                            let _ = InsertMenuW(
                                file_menu,
                                2 + i as u32,
                                MF_BYPOSITION | MF_STRING,
                                (IDM_FILE_RECENT_BASE + i as u32) as usize,
                                PCWSTR(text.as_ptr()),
                            );
                        }
                    }
                });
            }
        });
    }
}

/// Save window position to registry
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
                    ("ShowCmd", placement.showCmd),
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

        if let (Some(l), Some(t), Some(r), Some(b), Some(cmd)) = (left, top, right, bottom, show_cmd) {
            // Validate window is at least partially visible
            if r > l && b > t && (r - l) >= 100 && (b - t) >= 100 {
                let placement = WINDOWPLACEMENT {
                    length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
                    showCmd: cmd as u32,
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

/// Load a file into the WebView
fn load_file_into_webview(hwnd: HWND, file_path: &str) {
    // Read file
    let content = match fs::read_to_string(file_path) {
        Ok(c) => c,
        Err(e) => {
            let msg: Vec<u16> = format!("Failed to read file:\n{}\0", e)
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
    };

    // Convert to HTML
    let html_body = markdown::markdown_to_html(&content);
    let dark_mode = is_windows_dark_mode();
    let full_html = markdown::wrap_html(&html_body, dark_mode);

    // Navigate WebView
    CONTROLLER.with(|c| {
        if let Some(controller) = c.borrow().as_ref() {
            unsafe {
                if let Ok(webview) = controller.CoreWebView2() {
                    let html_wide = pwstr_from_str(&full_html);
                    let _ = webview.NavigateToString(PCWSTR(html_wide.as_ptr()));
                }
            }
        }
    });

    // Update window title
    let filename = Path::new(file_path)
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| file_path.to_string());
    let title: Vec<u16> = format!("{} - MDView\0", filename).encode_utf16().collect();
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
        // Load recent files from registry
        let recent = load_recent_files();
        RECENT_FILES.with(|files| {
            *files.borrow_mut() = recent;
        });

        // Register window class
        let class_name_wide: Vec<u16> = WINDOW_CLASS.encode_utf16().chain(std::iter::once(0)).collect();
        let hinstance = GetModuleHandleW(None)?;

        // Load application icon from resources
        let icon = LoadIconW(Some(hinstance.into()), PCWSTR(1 as *const u16)).ok();

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as _),
            lpszClassName: PCWSTR(class_name_wide.as_ptr()),
            hIcon: icon.unwrap_or_default(),
            hIconSm: icon.unwrap_or_default(),
            ..Default::default()
        };
        RegisterClassExW(&wc);

        // Create menu bar
        let menu = create_menu()?;
        MENU_HANDLE.with(|m| {
            *m.borrow_mut() = Some(menu);
        });

        // Create accelerator table
        let haccel = create_accelerators()?;
        ACCEL_HANDLE.with(|a| {
            *a.borrow_mut() = Some(haccel);
        });

        // Create main window
        let title_wide: Vec<u16> = format!("{} - MDView", title)
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        let hwnd = CreateWindowExW(
            WS_EX_ACCEPTFILES, // Enable drag & drop
            PCWSTR(class_name_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            1024,
            768,
            None,
            Some(menu),
            Some(hinstance.into()),
            None,
        )?;

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
            // Enable dark mode for menus via undocumented uxtheme API
            set_preferred_app_mode_dark();

            // Enable dark mode menu bar drawing
            DARK_MENU_BAR.with(|dmb| {
                dmb.borrow_mut().enable(hwnd);
            });

            // Enable non-client rendering
            let ncrp = DWMNCRP_ENABLED;
            let _ = DwmSetWindowAttribute(
                hwnd,
                DWMWA_NCRENDERING_POLICY,
                &ncrp as *const _ as *const std::ffi::c_void,
                std::mem::size_of_val(&ncrp) as u32,
            );

            // Enable immersive dark mode for title bar
            let use_dark_mode: i32 = 1;
            let _ = DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                &use_dark_mode as *const i32 as *const std::ffi::c_void,
                std::mem::size_of::<i32>() as u32,
            );
        }

        // Enable rounded corners on Windows 11
        let corner_preference = DWMWCP_ROUND.0;
        let _ = DwmSetWindowAttribute(
            hwnd,
            DWMWA_WINDOW_CORNER_PREFERENCE,
            &corner_preference as *const i32 as *const std::ffi::c_void,
            std::mem::size_of::<i32>() as u32,
        );

        // Enable Mica backdrop for modern look (Windows 11)
        let backdrop_type = DWMSBT_MAINWINDOW.0;
        let _ = DwmSetWindowAttribute(
            hwnd,
            DWMWA_SYSTEMBACKDROP_TYPE,
            &backdrop_type as *const i32 as *const std::ffi::c_void,
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
            let _ = SetWindowPos(
                hwnd,
                None,
                x,
                y,
                window_width,
                window_height,
                SWP_NOZORDER,
            );
        }

        // Show window
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

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
    static DARK_MENU_BAR: RefCell<dark_menu::DarkMenuBar> = RefCell::new(dark_menu::DarkMenuBar::new());
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
                                let _ = settings.SetAreDefaultContextMenusEnabled(true);
                                let _ = settings.SetIsStatusBarEnabled(false);
                            }

                            // Use NavigateToString for direct HTML loading
                            let html_wide = pwstr_from_str(&html_for_nav);
                            if let Err(e) = webview.NavigateToString(PCWSTR(html_wide.as_ptr())) {
                                eprintln!("NavigateToString failed: {:?}", e);
                            }

                            // Add message handler for Ctrl+click links
                            let handler = webview2_com::WebMessageReceivedEventHandler::create(
                                Box::new(|_webview, args| {
                                    if let Some(args) = args {
                                        let mut message_ptr: windows::core::PWSTR = windows::core::PWSTR::null();
                                        if args.WebMessageAsJson(&mut message_ptr).is_ok() && !message_ptr.is_null() {
                                            let msg_str = message_ptr.to_string().unwrap_or_default();
                                            windows::Win32::System::Com::CoTaskMemFree(Some(message_ptr.0 as *const _));
                                            // Parse simple JSON to extract URL
                                            if msg_str.contains("openLink") {
                                                if let Some(start) = msg_str.find("\"url\":\"") {
                                                    let url_start = start + 7;
                                                    if let Some(end) = msg_str[url_start..].find('"') {
                                                        let url = &msg_str[url_start..url_start + end];
                                                        // Open URL in default browser
                                                        let url_wide = U16CString::from_str(url).unwrap_or_default();
                                                        let open_wide = U16CString::from_str("open").unwrap_or_default();
                                                        ShellExecuteW(
                                                            None,
                                                            PCWSTR(open_wide.as_ptr()),
                                                            PCWSTR(url_wide.as_ptr()),
                                                            None,
                                                            None,
                                                            SW_SHOWNORMAL,
                                                        );
                                                    }
                                                }
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

    // Use TEMP folder for WebView2 (auto-cleaned by Windows)
    let user_data_folder = std::env::var("TEMP")
        .map(|p| format!("{}\\MDView_WebView2", p))
        .unwrap_or_else(|_| ".".to_string());
    let user_data_wide = pwstr_from_str(&user_data_folder);
    unsafe {
        let _ = CreateCoreWebView2EnvironmentWithOptions(None, PCWSTR(user_data_wide.as_ptr()), None, &env_handler);
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
    // Handle dark mode menu bar messages
    let dark_handled = DARK_MENU_BAR.with(|dmb| {
        dmb.borrow().handle_message(hwnd, msg, wparam, lparam)
    });
    if let Some(result) = dark_handled {
        return result;
    }

    match msg {
        WM_SIZE => {
            CONTROLLER.with(|c| {
                if let Some(controller) = c.borrow().as_ref() {
                    unsafe {
                        let mut rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut rect);
                        let _ = controller.SetBounds(rect);
                    }
                }
            });
            LRESULT(0)
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
                IDM_HELP_ABOUT => {
                    show_about_dialog(hwnd);
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
