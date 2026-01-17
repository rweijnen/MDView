use std::cell::RefCell;
use std::collections::HashMap;
use std::env;
use std::fs::OpenOptions;
use std::io::Write;
use std::rc::Rc;
use std::time::{Duration, Instant};

/// Simple debug logging to file (only in debug builds or when MDVIEW_DEBUG env is set)
#[allow(dead_code)]
fn log_debug(msg: &str) {
    if cfg!(debug_assertions) || env::var("MDVIEW_DEBUG").is_ok() {
        if let Ok(temp) = env::var("TEMP") {
            let log_path = format!("{}\\mdview_debug.log", temp);
            if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(&log_path) {
                let timestamp = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_secs())
                    .unwrap_or(0);
                let _ = writeln!(file, "[{}] {}", timestamp, msg);
            }
        }
    }
}
use webview2_com::Microsoft::Web::WebView2::Win32::*;
use webview2_com::{
    pwstr_from_str, CreateCoreWebView2ControllerCompletedHandler,
    CreateCoreWebView2EnvironmentCompletedHandler,
};
use widestring::U16CString;
use windows::core::PCWSTR;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::*;

const WINDOW_CLASS: &str = "MDViewWebView2Host";

// Thread-local storage for WebView2 controllers (COM objects are single-threaded)
thread_local! {
    static CONTROLLERS: RefCell<HashMap<isize, Rc<ICoreWebView2Controller>>> = RefCell::new(HashMap::new());
}

pub fn create_viewer(parent: HWND, html: &str) -> windows::core::Result<HWND> {
    unsafe {
        // Initialize COM if not already initialized (safe to call multiple times)
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

        // Register window class
        register_window_class()?;

        // Get parent client rect
        let mut rect = RECT::default();
        let _ = GetClientRect(parent, &mut rect);

        // Create host window
        let class_name = U16CString::from_str(WINDOW_CLASS).unwrap();
        let hinstance = GetModuleHandleW(None)?;
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            PCWSTR(class_name.as_ptr()),
            PCWSTR::null(),
            WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
            0,
            0,
            rect.right - rect.left,
            rect.bottom - rect.top,
            Some(parent),
            None,
            Some(hinstance.into()),
            None,
        )?;

        // Initialize WebView2 synchronously
        if let Err(e) = init_webview2_sync(hwnd, html) {
            let _ = DestroyWindow(hwnd);
            return Err(e);
        }

        Ok(hwnd)
    }
}

fn register_window_class() -> windows::core::Result<()> {
    unsafe {
        let class_name = U16CString::from_str(WINDOW_CLASS).unwrap();
        let hinstance = GetModuleHandleW(None)?;

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as _),
            lpszClassName: PCWSTR(class_name.as_ptr()),
            ..Default::default()
        };

        // Try to register, ignore if already registered
        let _ = RegisterClassExW(&wc);
        Ok(())
    }
}

fn init_webview2_sync(hwnd: HWND, html: &str) -> windows::core::Result<()> {
    log_debug(&format!("init_webview2_sync started, hwnd={:?}", hwnd.0));

    let html_owned = html.to_string();
    let controller_result: Rc<RefCell<Option<ICoreWebView2Controller>>> = Rc::new(RefCell::new(None));
    let error_result: Rc<RefCell<Option<windows::core::Error>>> = Rc::new(RefCell::new(None));
    let completed = Rc::new(RefCell::new(false));

    let controller_clone = controller_result.clone();
    let error_clone = error_result.clone();
    let completed_clone = completed.clone();

    // Create the environment first
    log_debug("Creating environment handler");
    let env_handler = CreateCoreWebView2EnvironmentCompletedHandler::create(Box::new(
        move |error_code, environment| {
            log_debug(&format!("Environment callback fired, error_code={:?}", error_code));
            let environment = match environment {
                Some(env) => env,
                None => {
                    log_debug("Environment is None, failing");
                    *error_clone.borrow_mut() = Some(windows::core::Error::from(E_FAIL));
                    *completed_clone.borrow_mut() = true;
                    return Ok(());
                }
            };
            log_debug("Environment created successfully");

            let controller_inner = controller_clone.clone();
            let error_inner = error_clone.clone();
            let completed_inner = completed_clone.clone();
            let html_for_nav = html_owned.clone();

            // Create the controller
            log_debug("Creating controller handler");
            let ctrl_handler = CreateCoreWebView2ControllerCompletedHandler::create(Box::new(
                move |error_code, controller| {
                    log_debug(&format!("Controller callback fired, error_code={:?}", error_code));
                    let controller = match controller {
                        Some(ctrl) => ctrl,
                        None => {
                            log_debug("Controller is None, failing");
                            *error_inner.borrow_mut() = Some(windows::core::Error::from(E_FAIL));
                            *completed_inner.borrow_mut() = true;
                            return Ok(());
                        }
                    };
                    log_debug("Controller created successfully");

                    unsafe {
                        // Make controller visible
                        let _ = controller.SetIsVisible(true);

                        // Set bounds
                        let mut rect = RECT::default();
                        let _ = GetClientRect(hwnd, &mut rect);
                        let _ = controller.SetBounds(rect);

                        // Get webview
                        if let Ok(webview) = controller.CoreWebView2() {
                            // Configure settings for a secure, read-only viewer
                            if let Ok(settings) = webview.Settings() {
                                let _ = settings.SetIsScriptEnabled(true); // needed for Ctrl+click and ESC handling
                                let _ = settings.SetAreDefaultContextMenusEnabled(false); // disable right-click menu
                                let _ = settings.SetAreDevToolsEnabled(false); // disable F12 dev tools
                                let _ = settings.SetIsStatusBarEnabled(false); // disable status bar
                                let _ = settings.SetIsBuiltInErrorPageEnabled(false); // disable error pages
                                let _ = settings.SetAreDefaultScriptDialogsEnabled(false); // disable alert/confirm/prompt
                                // Zoom control left enabled for accessibility (Ctrl+scroll)
                            }

                            // Navigate to HTML directly
                            let html_wide = pwstr_from_str(&html_for_nav);
                            let _ = webview.NavigateToString(PCWSTR(html_wide.as_ptr()));

                            // Add message handler for Ctrl+click links and ESC to close
                            let parent_hwnd = hwnd;
                            let handler = webview2_com::WebMessageReceivedEventHandler::create(
                                Box::new(move |_webview, args| {
                                    if let Some(args) = args {
                                        let mut message_ptr: windows::core::PWSTR = windows::core::PWSTR::null();
                                        if args.WebMessageAsJson(&mut message_ptr).is_ok() && !message_ptr.is_null() {
                                            let msg_str = message_ptr.to_string().unwrap_or_default();
                                            windows::Win32::System::Com::CoTaskMemFree(Some(message_ptr.0 as *const _));

                                            if msg_str.contains("\"close\"") {
                                                // Send WM_CLOSE to parent (TC lister window)
                                                if let Ok(parent) = GetParent(parent_hwnd) {
                                                    if !parent.is_invalid() {
                                                        let _ = PostMessageW(Some(parent), WM_CLOSE, WPARAM(0), LPARAM(0));
                                                    }
                                                }
                                            } else if msg_str.contains("openLink") {
                                                if let Some(start) = msg_str.find("\"url\":\"") {
                                                    let url_start = start + 7;
                                                    if let Some(end) = msg_str[url_start..].find('"') {
                                                        let url = &msg_str[url_start..url_start + end];
                                                        // Open URL in default browser
                                                        open_url_in_browser(url);
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

                        // Add accelerator key handler to pass F3, F5, F7 etc. to parent (TC)
                        let accel_parent = hwnd;
                        let accel_handler = webview2_com::AcceleratorKeyPressedEventHandler::create(
                            Box::new(move |_controller, args| {
                                if let Some(args) = args {
                                    use windows::Win32::UI::Input::KeyboardAndMouse::*;
                                    let mut key: u32 = 0;
                                    let mut key_event_kind: COREWEBVIEW2_KEY_EVENT_KIND = COREWEBVIEW2_KEY_EVENT_KIND::default();

                                    if args.VirtualKey(&mut key).is_ok() && args.KeyEventKind(&mut key_event_kind).is_ok() {
                                        // Only handle key down events
                                        if key_event_kind == COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN
                                            || key_event_kind == COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN
                                        {
                                            // Pass F3, F5, F7, N keys to parent (Total Commander)
                                            let pass_to_parent = matches!(
                                                VIRTUAL_KEY(key as u16),
                                                VK_F3 | VK_F5 | VK_F7 | VK_N
                                            );

                                            if pass_to_parent {
                                                // Mark as handled so WebView2 doesn't process it
                                                let _ = args.SetHandled(true);
                                                // Forward to parent window
                                                if let Ok(parent) = GetParent(accel_parent) {
                                                    if !parent.is_invalid() {
                                                        let _ = PostMessageW(
                                                            Some(parent),
                                                            WM_KEYDOWN,
                                                            WPARAM(key as usize),
                                                            LPARAM(0),
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                Ok(())
                            }),
                        );
                        let mut accel_token: i64 = 0;
                        let _ = controller.add_AcceleratorKeyPressed(&accel_handler, &mut accel_token);
                    }

                    log_debug("Storing controller and marking complete");
                    *controller_inner.borrow_mut() = Some(controller);
                    *completed_inner.borrow_mut() = true;
                    Ok(())
                },
            ));

            log_debug("Calling CreateCoreWebView2Controller");
            unsafe {
                let _ = environment.CreateCoreWebView2Controller(hwnd, &ctrl_handler);
            }
            log_debug("CreateCoreWebView2Controller returned");
            Ok(())
        },
    ));

    // Start the creation process - use TEMP folder (auto-cleaned by Windows)
    let user_data_folder = env::var("TEMP")
        .map(|p| format!("{}\\MDView_WebView2", p))
        .unwrap_or_else(|_| ".".to_string());
    log_debug(&format!("User data folder: {}", user_data_folder));
    let user_data_wide = pwstr_from_str(&user_data_folder);
    log_debug("Calling CreateCoreWebView2EnvironmentWithOptions");
    unsafe {
        let _ = CreateCoreWebView2EnvironmentWithOptions(None, PCWSTR(user_data_wide.as_ptr()), None, &env_handler);
    }
    log_debug("CreateCoreWebView2EnvironmentWithOptions returned, entering message loop");

    // Pump messages until completion (with timeout)
    let timeout = Duration::from_secs(30);
    let start = Instant::now();
    let mut loop_count = 0u32;
    unsafe {
        while !*completed.borrow() {
            loop_count += 1;

            // Check timeout
            if start.elapsed() > timeout {
                log_debug("TIMEOUT waiting for WebView2 initialization");
                *error_result.borrow_mut() = Some(windows::core::Error::from(E_FAIL));
                break;
            }

            // Log every 100 iterations
            if loop_count % 100 == 0 {
                log_debug(&format!("Message loop iteration {}, elapsed={:?}", loop_count, start.elapsed()));
            }

            let mut msg = MSG::default();
            // Use PeekMessage with timeout to avoid blocking forever
            if PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                if msg.message == WM_QUIT {
                    log_debug("Received WM_QUIT");
                    break;
                }
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            } else {
                // No message, sleep briefly to avoid spinning CPU
                std::thread::sleep(Duration::from_millis(10));
            }
        }
    }
    log_debug(&format!("Message loop exited, completed={}, loop_count={}", *completed.borrow(), loop_count));

    // Check for errors
    if let Some(err) = error_result.borrow_mut().take() {
        return Err(err);
    }

    // Store the controller
    if let Some(controller) = controller_result.borrow_mut().take() {
        CONTROLLERS.with(|c| {
            c.borrow_mut().insert(hwnd.0 as isize, Rc::new(controller));
        });
    }

    Ok(())
}

pub fn close_window(hwnd: HWND) {
    // Remove and close the controller
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow_mut().remove(&(hwnd.0 as isize)) {
            let _ = unsafe { controller.Close() };
        }
    });

    // Destroy the window
    let _ = unsafe { DestroyWindow(hwnd) };
}

pub fn execute_script(hwnd: HWND, script: &str) {
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow().get(&(hwnd.0 as isize)) {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                let script_wide = pwstr_from_str(script);
                let _ = unsafe { webview.ExecuteScript(PCWSTR(script_wide.as_ptr()), None) };
            }
        }
    });
}

fn resize_webview(hwnd: HWND) {
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow().get(&(hwnd.0 as isize)) {
            unsafe {
                let mut rect = RECT::default();
                let _ = GetClientRect(hwnd, &mut rect);
                let _ = controller.SetBounds(rect);
            }
        }
    });
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_SIZE => {
            resize_webview(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

fn open_url_in_browser(url: &str) {
    let url_wide = U16CString::from_str(url).unwrap_or_default();
    let open_wide = U16CString::from_str("open").unwrap_or_default();
    unsafe {
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
