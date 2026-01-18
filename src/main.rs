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
    DwmDefWindowProc, DwmExtendFrameIntoClientArea, DwmSetWindowAttribute, DWMNCRP_ENABLED,
    DWMWA_NCRENDERING_POLICY, DWMWA_USE_IMMERSIVE_DARK_MODE, DWMWA_WINDOW_CORNER_PREFERENCE,
    DWMWCP_ROUND, DWMWA_BORDER_COLOR,
};
use windows::Win32::UI::Controls::{MARGINS, DRAWITEMSTRUCT, MEASUREITEMSTRUCT, ODT_MENU, ODS_SELECTED, ODS_GRAYED};
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::GetCurrentThreadId;
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

// Custom window chrome base values (at 96 DPI)
const TITLE_BAR_HEIGHT_BASE: i32 = 40;  // Base height of custom title bar (matches Notepad)
const MENU_BAR_HEIGHT_BASE: i32 = 33;   // Base height of custom menu bar (matches Notepad)
const BUTTON_WIDTH_BASE: i32 = 46;      // Base width of window buttons

// Helper functions to get DPI-scaled dimensions
fn get_dpi_for_window(hwnd: HWND) -> i32 {
    unsafe {
        let hdc = GetDC(Some(hwnd));
        let dpi = GetDeviceCaps(Some(hdc), LOGPIXELSY);
        ReleaseDC(Some(hwnd), hdc);
        if dpi > 0 { dpi } else { 96 }
    }
}

fn scale_for_dpi(value: i32, dpi: i32) -> i32 {
    value * dpi / 96
}

fn get_title_bar_height(dpi: i32) -> i32 {
    scale_for_dpi(TITLE_BAR_HEIGHT_BASE, dpi)
}

fn get_menu_bar_height(dpi: i32) -> i32 {
    scale_for_dpi(MENU_BAR_HEIGHT_BASE, dpi)
}

fn get_nc_height(dpi: i32) -> i32 {
    get_title_bar_height(dpi) + get_menu_bar_height(dpi)
}

fn get_button_width(dpi: i32) -> i32 {
    scale_for_dpi(BUTTON_WIDTH_BASE, dpi)
}

// Dark mode colors (BGR format) - matching Notepad's style
// Active state (window has focus) - darker
const DARK_TITLEBAR_ACTIVE: u32 = 0x00000000;   // #000000 - title bar when active (full black)
pub const DARK_MENUBAR_ACTIVE: u32 = 0x00181818;    // #181818 - menu bar when active (between black and content)
// Inactive state (window doesn't have focus) - use previous active colors
const DARK_TITLEBAR_INACTIVE: u32 = 0x001E1E1E; // #1E1E1E - title bar when inactive
const DARK_MENUBAR_INACTIVE: u32 = 0x002D2D2D;  // #2D2D2D - menu bar when inactive
// Common colors
const DARK_HOVER_COLOR: u32 = 0x00404040;       // #404040 - hover highlight
const DARK_TEXT_COLOR: u32 = 0x00FFFFFF;        // White text

// Menu item IDs
const IDM_FILE_OPEN: u32 = 1001;
const IDM_FILE_EXIT: u32 = 1002;
const IDM_HELP_ABOUT: u32 = 2001;
const IDM_FILE_RECENT_BASE: u32 = 1100; // 1100-1109 for recent files

// Owner-drawn menu item data
#[derive(Clone)]
struct MenuItemData {
    id: u32,
    text: String,
    is_separator: bool,
}

// Storage for menu item data (needs to live as long as menu exists)
thread_local! {
    static MENU_ITEM_DATA: RefCell<Vec<Box<MenuItemData>>> = const { RefCell::new(Vec::new()) };
}

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

/// Allow dark mode for a specific window (needed for dark popup menus)
/// This calls AllowDarkModeForWindow (ordinal 133)
fn allow_dark_mode_for_window(hwnd: HWND) {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

    unsafe {
        let uxtheme: Vec<u16> = "uxtheme.dll\0".encode_utf16().collect();
        if let Ok(hmodule) = LoadLibraryW(PCWSTR(uxtheme.as_ptr())) {
            // Ordinal 133 = AllowDarkModeForWindow
            if let Some(allow_dark_mode) =
                GetProcAddress(hmodule, windows::core::PCSTR(133_usize as *const u8))
            {
                let func: extern "system" fn(isize, i32) -> i32 = std::mem::transmute(allow_dark_mode);
                func(hwnd.0 as isize, 1); // Allow dark mode for this window
            }
        }
    }
}

/// Refresh immersive color policy state (ordinal 104)
fn refresh_immersive_color_policy() {
    use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};

    unsafe {
        let uxtheme: Vec<u16> = "uxtheme.dll\0".encode_utf16().collect();
        if let Ok(hmodule) = LoadLibraryW(PCWSTR(uxtheme.as_ptr())) {
            // Ordinal 104 = RefreshImmersiveColorPolicyState
            if let Some(refresh) =
                GetProcAddress(hmodule, windows::core::PCSTR(104_usize as *const u8))
            {
                let func: extern "system" fn() = std::mem::transmute(refresh);
                func();
            }
        }
    }
}

/// Helper to create owner-drawn menu item data and return the pointer
fn create_menu_item_data(id: u32, text: &str, is_separator: bool) -> usize {
    let item_data = Box::new(MenuItemData {
        id,
        text: text.to_string(),
        is_separator,
    });
    let item_ptr = Box::into_raw(item_data) as usize;

    // Store pointer so we can free it later
    MENU_ITEM_DATA.with(|data| {
        data.borrow_mut().push(unsafe { Box::from_raw(item_ptr as *mut MenuItemData) });
    });

    // Get the pointer again (it's still valid, just also stored)
    MENU_ITEM_DATA.with(|data| {
        let items = data.borrow();
        items.last().map(|b| b.as_ref() as *const MenuItemData as usize).unwrap_or(0)
    })
}

/// Helper to add an owner-drawn menu item (append)
fn add_owner_drawn_item(menu: HMENU, id: u32, text: &str, is_separator: bool) -> windows::core::Result<()> {
    unsafe {
        let ptr = create_menu_item_data(id, text, is_separator);

        if is_separator {
            AppendMenuW(menu, MF_OWNERDRAW | MF_SEPARATOR, 0, PCWSTR(ptr as *const u16))?;
        } else {
            AppendMenuW(menu, MF_OWNERDRAW, id as usize, PCWSTR(ptr as *const u16))?;
        }
        Ok(())
    }
}

/// Helper to insert an owner-drawn menu item at a position
fn insert_owner_drawn_item(menu: HMENU, position: u32, id: u32, text: &str, is_grayed: bool) -> windows::core::Result<()> {
    unsafe {
        let ptr = create_menu_item_data(id, text, false);
        let flags = if is_grayed {
            MF_BYPOSITION | MF_OWNERDRAW | MF_GRAYED
        } else {
            MF_BYPOSITION | MF_OWNERDRAW
        };
        InsertMenuW(menu, position, flags, id as usize, PCWSTR(ptr as *const u16))?;
        Ok(())
    }
}

/// Configure popup menu with dark theme background
fn configure_dark_popup_menu(menu: HMENU) {
    unsafe {
        // Create dark background brush for popup menu
        let dark_brush = CreateSolidBrush(COLORREF(0x002D2D2D));

        let mut mi = MENUINFO {
            cbSize: std::mem::size_of::<MENUINFO>() as u32,
            fMask: MIM_BACKGROUND | MIM_APPLYTOSUBMENUS,
            dwStyle: MENUINFO_STYLE::default(),
            cyMax: 0,
            hbrBack: dark_brush,
            dwContextHelpID: 0,
            dwMenuData: 0,
        };
        let _ = SetMenuInfo(menu, &mi);
    }
}

/// Create the application menu bar
fn create_menu() -> windows::core::Result<HMENU> {
    unsafe {
        // Clear any existing menu item data
        MENU_ITEM_DATA.with(|data| data.borrow_mut().clear());

        let menu_bar = CreateMenu()?;
        let file_menu = CreatePopupMenu()?;
        let help_menu = CreatePopupMenu()?;

        // Configure popup menus with dark background
        configure_dark_popup_menu(file_menu);
        configure_dark_popup_menu(help_menu);

        // File menu items (owner-drawn)
        add_owner_drawn_item(file_menu, IDM_FILE_OPEN, "&Open\tCtrl+O", false)?;
        add_owner_drawn_item(file_menu, 0, "", true)?; // Separator
        add_owner_drawn_item(file_menu, 0, "Recent Files", false)?; // Placeholder
        add_owner_drawn_item(file_menu, 0, "", true)?; // Separator
        add_owner_drawn_item(file_menu, IDM_FILE_EXIT, "E&xit", false)?;

        // Help menu items (owner-drawn)
        add_owner_drawn_item(help_menu, IDM_HELP_ABOUT, "&About MDView", false)?;

        // Add submenus to menu bar (these stay as MF_POPUP)
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
                        // Show disabled placeholder (owner-drawn)
                        let _ = insert_owner_drawn_item(file_menu, 2, 0, "(No Recent Files)", true);
                    } else {
                        // Add recent files (owner-drawn)
                        for (i, path) in files.iter().enumerate() {
                            let filename = Path::new(path)
                                .file_name()
                                .map(|n| n.to_string_lossy().to_string())
                                .unwrap_or_else(|| path.clone());
                            let text = format!("&{} {}", i + 1, filename);
                            let _ = insert_owner_drawn_item(
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
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as _), // Standard window background
            lpszClassName: PCWSTR(class_name_wide.as_ptr()),
            hIcon: icon.unwrap_or_default(),
            hIconSm: icon.unwrap_or_default(),
            ..Default::default()
        };
        RegisterClassExW(&wc);

        // Create menu bar (stored later, not attached to window)
        let menu = create_menu()?;

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
            None, // No standard menu - we'll draw our own
            Some(hinstance.into()),
            None,
        )?;

        // Store menu handle for manual popup display
        // (We don't use SetMenu because our custom NCCALCSIZE conflicts with Windows' menu bar)
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
            // Refresh immersive color policy to apply dark mode
            refresh_immersive_color_policy();

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

            // Enable Mica backdrop for Windows 11 (attribute 38)
            // DWMSBT_MAINWINDOW = 2 (Mica), DWMSBT_TABBEDWINDOW = 4 (Mica Alt)
            const DWMSBT_MAINWINDOW: i32 = 2; // Mica
            let _ = DwmSetWindowAttribute(
                hwnd,
                windows::Win32::Graphics::Dwm::DWMWINDOWATTRIBUTE(38), // DWMWA_SYSTEMBACKDROP_TYPE
                &DWMSBT_MAINWINDOW as *const i32 as *const std::ffi::c_void,
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

        // Extend frame into client area for custom title bar and menu bar
        // We draw both ourselves to ensure color consistency
        let dpi = get_dpi_for_window(hwnd);
        let nc_height = get_nc_height(dpi);
        let margins = MARGINS {
            cxLeftWidth: 0,
            cxRightWidth: 0,
            cyTopHeight: nc_height, // Title bar + menu bar height (DPI-scaled)
            cyBottomHeight: 0,
        };
        let _ = DwmExtendFrameIntoClientArea(hwnd, &margins);

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
    static WINDOW_ACTIVE: RefCell<bool> = const { RefCell::new(true) };
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
    // Let DWM handle caption button interactions (hover, click effects)
    let mut dwm_result = LRESULT(0);
    if DwmDefWindowProc(hwnd, msg, wparam, lparam, &mut dwm_result).as_bool() {
        return dwm_result;
    }

    // Handle dark mode menu bar messages
    let dark_handled = DARK_MENU_BAR.with(|dmb| {
        dmb.borrow().handle_message(hwnd, msg, wparam, lparam)
    });
    if let Some(result) = dark_handled {
        return result;
    }

    match msg {
        WM_SIZE => {
            // Adjust WebView bounds to account for custom title bar and menu bar
            let dpi = get_dpi_for_window(hwnd);
            let nc_height = get_nc_height(dpi);
            CONTROLLER.with(|c| {
                if let Some(controller) = c.borrow().as_ref() {
                    unsafe {
                        let mut rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut rect);
                        // Offset for custom NC area (DPI-scaled)
                        rect.top += nc_height;
                        let _ = controller.SetBounds(rect);
                    }
                }
            });
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // Track window active state for title bar color
            let active = (wparam.0 & 0xFFFF) != 0; // WA_INACTIVE = 0
            WINDOW_ACTIVE.with(|a| *a.borrow_mut() = active);

            // Repaint the NC area to update colors
            unsafe {
                let dpi = get_dpi_for_window(hwnd);
                let nc_height = get_nc_height(dpi);
                let mut rc = RECT::default();
                let _ = GetClientRect(hwnd, &mut rc);
                rc.bottom = nc_height;
                let _ = InvalidateRect(Some(hwnd), Some(&rc), false);
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_ERASEBKGND => {
            // Fill title bar area with dark color to prevent white flash
            // Mica will paint over this for active windows
            unsafe {
                let hdc = HDC(wparam.0 as *mut _);
                let dpi = get_dpi_for_window(hwnd);
                let title_bar_height = get_title_bar_height(dpi);

                let mut rc = RECT::default();
                let _ = GetClientRect(hwnd, &mut rc);

                // Fill title bar area with dark color
                let rc_titlebar = RECT {
                    left: 0,
                    top: 0,
                    right: rc.right,
                    bottom: title_bar_height,
                };
                let brush = CreateSolidBrush(COLORREF(DARK_TITLEBAR_ACTIVE)); // Use black
                FillRect(hdc, &rc_titlebar, brush);
                let _ = DeleteObject(brush.into());
            }
            LRESULT(1) // Return non-zero to indicate we handled it
        }
        WM_NCCALCSIZE => {
            // Custom NC area: extend client area to include title bar and menu bar
            if wparam.0 != 0 {
                let params = lparam.0 as *mut NCCALCSIZE_PARAMS;
                unsafe {
                    let original_top = (*params).rgrc[0].top;

                    // Let Windows calculate default frame
                    let result = DefWindowProcW(hwnd, msg, wparam, lparam);
                    if result.0 != 0 {
                        return result;
                    }

                    // Restore original top to remove standard title bar
                    (*params).rgrc[0].top = original_top;
                }
            }
            LRESULT(0)
        }
        WM_NCHITTEST => {
            // Custom NC area: handle hit testing
            unsafe {
                let dpi = get_dpi_for_window(hwnd);
                let title_bar_height = get_title_bar_height(dpi);
                let nc_height = get_nc_height(dpi);
                let button_width = get_button_width(dpi);

                let pt = POINT {
                    x: (lparam.0 & 0xFFFF) as i16 as i32,
                    y: ((lparam.0 >> 16) & 0xFFFF) as i16 as i32,
                };

                // Get window rect
                let mut rc = RECT::default();
                let _ = GetWindowRect(hwnd, &mut rc);

                // Check if in resize border (top 8 pixels when not maximized)
                let mut wp = WINDOWPLACEMENT::default();
                wp.length = std::mem::size_of::<WINDOWPLACEMENT>() as u32;
                let _ = GetWindowPlacement(hwnd, &mut wp);
                let is_maximized = wp.showCmd == SW_SHOWMAXIMIZED.0 as u32;

                let resize_border = scale_for_dpi(8, dpi);
                if !is_maximized && pt.y < rc.top + resize_border {
                    return LRESULT(HTTOP as isize);
                }

                // Check if in title bar area
                if pt.y < rc.top + title_bar_height {
                    // Check for window buttons area (right side)
                    if pt.x > rc.right - button_width {
                        return LRESULT(HTCLOSE as isize);
                    }
                    if pt.x > rc.right - button_width * 2 {
                        return LRESULT(HTMAXBUTTON as isize);
                    }
                    if pt.x > rc.right - button_width * 3 {
                        return LRESULT(HTMINBUTTON as isize);
                    }
                    return LRESULT(HTCAPTION as isize);
                }

                // Check if in menu bar area
                if pt.y < rc.top + nc_height {
                    return LRESULT(HTCLIENT as isize);
                }

                // Let default handling take over
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_PAINT => {
            // Custom NC area: draw title bar and menu bar
            unsafe {
                let mut ps = PAINTSTRUCT::default();
                let hdc = BeginPaint(hwnd, &mut ps);

                // Get DPI for proper scaling
                let dpi = GetDeviceCaps(Some(hdc), LOGPIXELSY);
                let title_bar_height = get_title_bar_height(dpi);
                let menu_bar_height = get_menu_bar_height(dpi);
                let nc_height = title_bar_height + menu_bar_height;

                let mut rc_client = RECT::default();
                let _ = GetClientRect(hwnd, &mut rc_client);

                // Only draw if painting in our NC area
                if ps.rcPaint.top < nc_height {
                    // Get active state for color selection
                    let is_active = WINDOW_ACTIVE.with(|a| *a.borrow());
                    let menubar_color = if is_active { DARK_MENUBAR_ACTIVE } else { DARK_MENUBAR_INACTIVE };

                    let menubar_brush = CreateSolidBrush(COLORREF(menubar_color));

                    // Always paint title bar black - Mica blends with it for active windows
                    let titlebar_brush = CreateSolidBrush(COLORREF(DARK_TITLEBAR_ACTIVE));
                    let rc_titlebar = RECT {
                        left: 0,
                        top: 0,
                        right: rc_client.right,
                        bottom: title_bar_height,
                    };
                    FillRect(hdc, &rc_titlebar, titlebar_brush);
                    let _ = DeleteObject(titlebar_brush.into());

                    // Draw application icon in title bar (DPI-scaled)
                    let icon_size = scale_for_dpi(16, dpi);
                    let icon_x = scale_for_dpi(12, dpi);
                    let icon_y = (title_bar_height - icon_size) / 2;
                    if let Ok(hmodule) = GetModuleHandleW(None) {
                        let hinstance = HINSTANCE(hmodule.0);
                        if let Ok(hicon) = LoadIconW(Some(hinstance), PCWSTR(1 as *const u16)) {
                            let _ = DrawIconEx(
                                hdc,
                                icon_x,
                                icon_y,
                                hicon,
                                icon_size,
                                icon_size,
                                0,
                                None,
                                DI_NORMAL,
                            );
                        }
                    }

                    // Draw menu bar background (darker)
                    let rc_menubar = RECT {
                        left: 0,
                        top: title_bar_height,
                        right: rc_client.right,
                        bottom: nc_height,
                    };
                    FillRect(hdc, &rc_menubar, menubar_brush);

                    // Create font for title (Segoe UI, 12pt) - DPI aware
                    // Formula: height = -(pointSize * dpi / 72)
                    let title_font_height = -(12 * dpi / 72);
                    let font_name: Vec<u16> = "Segoe UI\0".encode_utf16().collect();
                    let title_font = CreateFontW(
                        title_font_height,
                        0, 0, 0,
                        FW_NORMAL.0 as i32,
                        0, 0, 0,
                        DEFAULT_CHARSET,
                        OUT_DEFAULT_PRECIS,
                        CLIP_DEFAULT_PRECIS,
                        CLEARTYPE_QUALITY,
                        DEFAULT_PITCH.0 as u32,
                        PCWSTR(font_name.as_ptr()),
                    );

                    // Draw window title
                    let mut title_buf = [0u16; 256];
                    let title_len = GetWindowTextW(hwnd, &mut title_buf);
                    if title_len > 0 {
                        let old_font = SelectObject(hdc, title_font.into());
                        let old_bk = SetBkMode(hdc, TRANSPARENT);
                        let old_color = SetTextColor(hdc, COLORREF(DARK_TEXT_COLOR));

                        let title_left = scale_for_dpi(36, dpi); // After icon (12 + 16 + 8)
                        let mut rc_title = RECT {
                            left: title_left,
                            top: 0,
                            right: rc_client.right - scale_for_dpi(150, dpi),
                            bottom: title_bar_height,
                        };

                        DrawTextW(
                            hdc,
                            &mut title_buf[..title_len as usize],
                            &mut rc_title,
                            DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS,
                        );

                        SetTextColor(hdc, old_color);
                        SetBkMode(hdc, BACKGROUND_MODE(old_bk as u32));
                        SelectObject(hdc, old_font);
                    }

                    // Create font for menu items (Segoe UI, 11pt) - DPI aware
                    let menu_font_height = -(11 * dpi / 72);
                    let menu_font = CreateFontW(
                        menu_font_height,
                        0, 0, 0,
                        FW_NORMAL.0 as i32,
                        0, 0, 0,
                        DEFAULT_CHARSET,
                        OUT_DEFAULT_PRECIS,
                        CLIP_DEFAULT_PRECIS,
                        CLEARTYPE_QUALITY,
                        DEFAULT_PITCH.0 as u32,
                        PCWSTR(font_name.as_ptr()),
                    );

                    let old_font = SelectObject(hdc, menu_font.into());
                    let old_bk = SetBkMode(hdc, TRANSPARENT);
                    let old_color = SetTextColor(hdc, COLORREF(DARK_TEXT_COLOR));

                    // Draw menu items with Notepad-style spacing (DPI-aware)
                    let menu_item_width = 56 * dpi / 96; // Width per menu item, scaled for DPI
                    let menu_padding = 0;                // No padding, flush left like Notepad

                    // File menu
                    let file_text: Vec<u16> = "File".encode_utf16().collect();
                    let mut rc_file = RECT {
                        left: menu_padding,
                        top: title_bar_height,
                        right: menu_padding + menu_item_width,
                        bottom: nc_height,
                    };
                    DrawTextW(hdc, &mut file_text.clone(), &mut rc_file, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

                    // Help menu
                    let help_text: Vec<u16> = "Help".encode_utf16().collect();
                    let mut rc_help = RECT {
                        left: menu_padding + menu_item_width,
                        top: title_bar_height,
                        right: menu_padding + menu_item_width * 2,
                        bottom: nc_height,
                    };
                    DrawTextW(hdc, &mut help_text.clone(), &mut rc_help, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

                    SetTextColor(hdc, old_color);
                    SetBkMode(hdc, BACKGROUND_MODE(old_bk as u32));
                    SelectObject(hdc, old_font);

                    // Clean up GDI objects
                    let _ = DeleteObject(title_font.into());
                    let _ = DeleteObject(menu_font.into());
                    let _ = DeleteObject(menubar_brush.into());
                }

                let _ = EndPaint(hwnd, &ps);
            }
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            // Handle menu clicks manually (since we can't use SetMenu with our custom NCCALCSIZE)
            unsafe {
                let x = (lparam.0 & 0xFFFF) as i16 as i32;
                let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as i32;

                // Get DPI-scaled dimensions
                let dpi = get_dpi_for_window(hwnd);
                let title_bar_height = get_title_bar_height(dpi);
                let nc_height = get_nc_height(dpi);
                let menu_item_width = scale_for_dpi(56, dpi);

                // Check if click is in menu bar area
                if y >= title_bar_height && y < nc_height {
                    MENU_HANDLE.with(|menu| {
                        if let Some(hmenu) = menu.borrow().as_ref() {
                            // Determine which menu was clicked
                            let (popup, menu_x) = if x >= 0 && x < menu_item_width {
                                // File menu
                                (GetSubMenu(*hmenu, 0), 0)
                            } else if x >= menu_item_width && x < menu_item_width * 2 {
                                // Help menu
                                (GetSubMenu(*hmenu, 1), menu_item_width)
                            } else {
                                (HMENU::default(), 0)
                            };

                            if !popup.is_invalid() {
                                let mut pt = POINT { x: menu_x, y: nc_height };
                                let _ = ClientToScreen(hwnd, &mut pt);

                                let _ = TrackPopupMenu(
                                    popup,
                                    TPM_LEFTALIGN | TPM_TOPALIGN,
                                    pt.x,
                                    pt.y,
                                    Some(0),
                                    hwnd,
                                    None,
                                );
                            }
                        }
                    });
                    return LRESULT(0);
                }
            }
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_MEASUREITEM => {
            // Owner-drawn menu: provide item dimensions (DPI-aware)
            unsafe {
                let mis = lparam.0 as *mut MEASUREITEMSTRUCT;
                if (*mis).CtlType == ODT_MENU {
                    let item_data = (*mis).itemData as *const MenuItemData;
                    if !item_data.is_null() {
                        let data = &*item_data;
                        let dpi = get_dpi_for_window(hwnd);
                        let scale = dpi as f32 / 96.0;

                        if data.is_separator {
                            // Separator: thin line (DPI-scaled)
                            (*mis).itemHeight = (9.0 * scale) as u32;
                            (*mis).itemWidth = 0;
                        } else {
                            // Regular item: match Notepad's ~36px at 96 DPI
                            (*mis).itemHeight = (36.0 * scale) as u32;
                            (*mis).itemWidth = (250.0 * scale) as u32;
                        }
                    }
                }
            }
            LRESULT(1) // TRUE = we handled it
        }
        WM_DRAWITEM => {
            // Owner-drawn menu: draw the item
            unsafe {
                let dis = lparam.0 as *const DRAWITEMSTRUCT;
                if (*dis).CtlType == ODT_MENU {
                    let item_data = (*dis).itemData as *const MenuItemData;
                    if !item_data.is_null() {
                        let data = &*item_data;
                        let hdc = (*dis).hDC;
                        let rc = (*dis).rcItem;
                        let is_selected = ((*dis).itemState.0 & ODS_SELECTED.0) != 0;
                        let is_disabled = ((*dis).itemState.0 & ODS_GRAYED.0) != 0;

                        // Background color - dark gray, lighter when selected
                        let bg_color = if is_selected {
                            COLORREF(0x00404040) // Lighter for hover
                        } else {
                            COLORREF(0x002D2D2D) // Dark background like Notepad
                        };
                        let brush = CreateSolidBrush(bg_color);
                        FillRect(hdc, &rc, brush);
                        let _ = DeleteObject(brush.into());

                        if data.is_separator {
                            // Draw separator line (DPI-aware)
                            let dpi = get_dpi_for_window(hwnd);
                            let scale = dpi as f32 / 96.0;
                            let margin = (12.0 * scale) as i32;
                            let sep_brush = CreateSolidBrush(COLORREF(0x00505050));
                            let sep_rc = RECT {
                                left: rc.left + margin,
                                top: (rc.top + rc.bottom) / 2,
                                right: rc.right - margin,
                                bottom: (rc.top + rc.bottom) / 2 + 1,
                            };
                            FillRect(hdc, &sep_rc, sep_brush);
                            let _ = DeleteObject(sep_brush.into());
                        } else {
                            // Draw text (DPI-aware)
                            let dpi = get_dpi_for_window(hwnd);
                            let scale = dpi as f32 / 96.0;

                            let text_color = if is_disabled {
                                COLORREF(0x00808080) // Gray for disabled
                            } else {
                                COLORREF(0x00FFFFFF) // White
                            };

                            let old_bk = SetBkMode(hdc, TRANSPARENT);
                            let old_color = SetTextColor(hdc, text_color);

                            // Create font (DPI-scaled, ~12pt at 96 DPI)
                            let font_height = (-14.0 * scale) as i32;
                            let font = CreateFontW(
                                font_height, 0, 0, 0,
                                FW_NORMAL.0 as i32,
                                0, 0, 0,
                                DEFAULT_CHARSET,
                                OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY,
                                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                                windows::core::w!("Segoe UI")
                            );
                            let old_font = SelectObject(hdc, font.into());

                            // Split text and shortcut (separated by \t)
                            let parts: Vec<&str> = data.text.split('\t').collect();
                            let label = parts.first().unwrap_or(&"");
                            let shortcut = parts.get(1);

                            // Draw label (left aligned with DPI-scaled padding)
                            let padding = (16.0 * scale) as i32;
                            let mut text_rc = RECT {
                                left: rc.left + padding,
                                top: rc.top,
                                right: rc.right - padding,
                                bottom: rc.bottom,
                            };
                            let mut label_wide: Vec<u16> = label.encode_utf16().collect();
                            DrawTextW(hdc, &mut label_wide, &mut text_rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

                            // Draw shortcut (right aligned, slightly dimmer)
                            if let Some(sc) = shortcut {
                                let shortcut_color = COLORREF(0x00A0A0A0); // Dimmer white
                                SetTextColor(hdc, shortcut_color);
                                let mut sc_wide: Vec<u16> = sc.encode_utf16().collect();
                                DrawTextW(hdc, &mut sc_wide, &mut text_rc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
                            }

                            SelectObject(hdc, old_font);
                            let _ = DeleteObject(font.into());
                            SetTextColor(hdc, old_color);
                            SetBkMode(hdc, BACKGROUND_MODE(old_bk as u32));
                        }
                    }
                }
            }
            LRESULT(1) // TRUE = we handled it
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
