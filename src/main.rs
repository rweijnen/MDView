mod markdown;
mod terminal;

use std::cell::RefCell;
use std::env;
use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;
use std::rc::Rc;

fn print_usage() {
    eprintln!("MDView - Markdown Viewer v{}", env!("CARGO_PKG_VERSION"));
    eprintln!("Copyright 2026 Remko Weijnen - Mozilla Public License 2.0");
    eprintln!("https://github.com/rweijnen/MDView");
    eprintln!();
    eprintln!("Usage: mdview [OPTIONS] [FILE]");
    eprintln!();
    eprintln!("Options:");
    eprintln!("  --gui        Open in GUI window (default if file specified interactively)");
    eprintln!("  --term       Output with terminal colors/formatting (default for piped output)");
    eprintln!("  --html       Output full HTML document to stdout");
    eprintln!("  --body       Output HTML body only (no wrapper)");
    eprintln!("  --text       Output plain text (no formatting)");
    eprintln!("  --dark       Use dark mode colors (GUI and HTML only)");
    eprintln!("  -h, --help   Show this help message");
    eprintln!();
    eprintln!("If no FILE is specified, reads from stdin (CLI mode only).");
    eprintln!();
    eprintln!("Terminal output features (Windows Terminal, modern terminals):");
    eprintln!("  - Clickable hyperlinks (OSC 8)");
    eprintln!("  - True color syntax highlighting");
    eprintln!("  - Unicode box drawing for tables");
    eprintln!("  - Bold, italic, strikethrough formatting");
    eprintln!();
    eprintln!("Examples:");
    eprintln!("  mdview README.md              # Open in GUI window");
    eprintln!("  mdview --term README.md       # Output with terminal colors");
    eprintln!("  cat doc.md | mdview           # Piped input, terminal output");
    eprintln!("  mdview --html README.md       # Output HTML to stdout");
}

#[derive(Default)]
struct Options {
    gui_mode: bool,
    terminal_mode: bool,
    html_full: bool,
    html_body: bool,
    plain_text: bool,
    dark_mode: bool,
    file_path: Option<String>,
}

fn parse_args() -> Result<Options, String> {
    let args: Vec<String> = env::args().skip(1).collect();
    let mut opts = Options::default();

    for arg in args {
        match arg.as_str() {
            "-h" | "--help" => {
                print_usage();
                std::process::exit(0);
            }
            "--gui" => opts.gui_mode = true,
            "--term" | "--terminal" => opts.terminal_mode = true,
            "--html" => opts.html_full = true,
            "--body" => opts.html_body = true,
            "--text" => opts.plain_text = true,
            "--dark" => opts.dark_mode = true,
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

    // Default behavior: GUI if file specified and running interactively, otherwise terminal output
    if !opts.gui_mode && cli_format_count == 0 {
        if opts.file_path.is_some() && atty::is(atty::Stream::Stdout) {
            opts.gui_mode = true;
        } else {
            opts.terminal_mode = true; // Default to terminal output for piped mode
        }
    }

    Ok(opts)
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
    let opts = match parse_args() {
        Ok(o) => o,
        Err(e) => {
            eprintln!("Error: {}", e);
            eprintln!("Use --help for usage information.");
            std::process::exit(1);
        }
    };

    // If no file and no stdin data expected in GUI mode, show help
    if opts.file_path.is_none() && opts.gui_mode {
        eprintln!("Error: GUI mode requires a file argument.");
        print_usage();
        std::process::exit(1);
    }

    // If no file and running interactively in CLI mode, show help
    if opts.file_path.is_none() && !opts.gui_mode && atty::is(atty::Stream::Stdin) {
        print_usage();
        std::process::exit(0);
    }

    let markdown_content = match read_input(opts.file_path.as_deref()) {
        Ok(content) => content,
        Err(e) => {
            eprintln!("Error reading input: {}", e);
            std::process::exit(1);
        }
    };

    if opts.gui_mode {
        // GUI mode - open window with WebView2
        let html_body = markdown::markdown_to_html(&markdown_content);
        let full_html = markdown::wrap_html(&html_body, opts.dark_mode);
        let title = opts
            .file_path
            .as_ref()
            .map(|p| Path::new(p).file_name().unwrap_or_default().to_string_lossy().to_string())
            .unwrap_or_else(|| "MDView".to_string());

        if let Err(e) = run_gui(&title, &full_html) {
            eprintln!("Error: {}", e);
            std::process::exit(1);
        }
    } else {
        // CLI mode - output to stdout
        let output = if opts.terminal_mode {
            let caps = terminal::TerminalCaps::detect();
            terminal::render_to_terminal(&markdown_content, &caps)
        } else if opts.plain_text {
            markdown::markdown_to_plain_text(&markdown_content)
        } else if opts.html_body {
            markdown::markdown_to_html(&markdown_content)
        } else {
            let html_body = markdown::markdown_to_html(&markdown_content);
            markdown::wrap_html(&html_body, opts.dark_mode)
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
use widestring::U16CString;
use windows::core::PCWSTR;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::*;

const WINDOW_CLASS: &str = "MDViewWindow";

fn run_gui(title: &str, html: &str) -> windows::core::Result<()> {
    unsafe {
        // Initialize COM for this thread (required for WebView2)
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(windows::core::Error::from(hr));
        }

        let result = run_gui_inner(title, html);

        CoUninitialize();
        result
    }
}

fn run_gui_inner(title: &str, html: &str) -> windows::core::Result<()> {
    unsafe {
        // Register window class
        let class_name_wide: Vec<u16> = WINDOW_CLASS.encode_utf16().chain(std::iter::once(0)).collect();
        let hinstance = GetModuleHandleW(None)?;

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as _),
            lpszClassName: PCWSTR(class_name_wide.as_ptr()),
            ..Default::default()
        };
        RegisterClassExW(&wc);

        // Create main window
        let title_wide: Vec<u16> = format!("{} - MDView", title)
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            PCWSTR(class_name_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            1024,
            768,
            None,
            None,
            Some(hinstance.into()),
            None,
        )?;

        // Initialize WebView2
        init_webview2_gui(hwnd, html)?;

        // Show window
        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

        // Message loop
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        Ok(())
    }
}

thread_local! {
    static CONTROLLER: RefCell<Option<ICoreWebView2Controller>> = const { RefCell::new(None) };
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
        WM_DESTROY => {
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
