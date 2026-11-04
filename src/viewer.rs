use std::cell::RefCell;
use std::collections::HashMap;
use std::env;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::Path;
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
    CreateCoreWebView2ControllerCompletedHandler, CreateCoreWebView2EnvironmentCompletedHandler,
    pwstr_from_str,
};
use widestring::U16CString;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{SHCreateMemStream, ShellExecuteW};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::PCWSTR;

const WINDOW_CLASS_PREFIX: &str = "MDViewWebView2Host";

fn window_class_name() -> U16CString {
    let unique = format!("{}_{}", WINDOW_CLASS_PREFIX, window_proc as *const () as usize);
    U16CString::from_str(&unique).unwrap()
}

// Thread-local storage for WebView2 controllers (COM objects are single-threaded)
thread_local! {
    static CONTROLLERS: RefCell<HashMap<isize, Rc<ICoreWebView2Controller>>> = RefCell::new(HashMap::new());
    static CURRENT_FILES: RefCell<HashMap<isize, String>> = RefCell::new(HashMap::new());
    static DARK_MODES: RefCell<HashMap<isize, bool>> = RefCell::new(HashMap::new());
}

pub fn create_viewer(
    parent: HWND,
    html: &str,
    file_path: Option<&str>,
    dark_mode: bool,
) -> windows::core::Result<HWND> {
    unsafe {
        // Initialize COM if not already initialized (safe to call multiple times)
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);

        // Register window class
        register_window_class()?;

        // Get parent client rect
        let mut rect = RECT::default();
        let _ = GetClientRect(parent, &mut rect);

        // Create host window
        let class_name = window_class_name();
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

        if let Some(file_path) = file_path {
            CURRENT_FILES.with(|f| {
                f.borrow_mut()
                    .insert(hwnd.0 as isize, file_path.to_string());
            });
        }
        DARK_MODES.with(|m| {
            m.borrow_mut().insert(hwnd.0 as isize, dark_mode);
        });

        // Initialize WebView2 synchronously
        if let Err(e) = init_webview2_sync(hwnd, html) {
            CURRENT_FILES.with(|f| {
                f.borrow_mut().remove(&(hwnd.0 as isize));
            });
            DARK_MODES.with(|m| {
                m.borrow_mut().remove(&(hwnd.0 as isize));
            });
            let _ = DestroyWindow(hwnd);
            return Err(e);
        }

        Ok(hwnd)
    }
}

fn register_window_class() -> windows::core::Result<()> {
    unsafe {
        let class_name = window_class_name();
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

fn absolute_path(path: &Path) -> Option<std::path::PathBuf> {
    if path.is_absolute() {
        Some(path.to_path_buf())
    } else {
        std::env::current_dir().ok().map(|cwd| cwd.join(path))
    }
}

fn get_path_root(path: &Path) -> Option<std::path::PathBuf> {
    use std::path::Component;

    let absolute = absolute_path(path)?;
    let mut components = absolute.components();

    match components.next()? {
        Component::Prefix(prefix) => {
            if !matches!(components.next(), Some(Component::RootDir)) {
                return None;
            }

            let mut root = std::path::PathBuf::from(prefix.as_os_str());
            root.push(std::path::MAIN_SEPARATOR.to_string());
            Some(root)
        }
        Component::RootDir => Some(std::path::PathBuf::from(
            std::path::MAIN_SEPARATOR.to_string(),
        )),
        _ => None,
    }
}

fn path_relative_to_root(path: &Path) -> Option<String> {
    use std::path::Component;

    let absolute = absolute_path(path)?;
    let mut saw_root = false;
    let mut parts = Vec::new();

    for component in absolute.components() {
        match component {
            Component::Prefix(_) => {}
            Component::RootDir => saw_root = true,
            Component::Normal(part) => parts.push(part.to_string_lossy().into_owned()),
            Component::CurDir => {}
            Component::ParentDir => {
                let _ = parts.pop();
            }
        }
    }

    if !saw_root {
        return None;
    }

    Some(parts.join("/"))
}

fn percent_encode_path(path: &str) -> String {
    let mut encoded = String::with_capacity(path.len());

    for byte in path.bytes() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' | b'/' => {
                encoded.push(byte as char)
            }
            _ => encoded.push_str(&format!("%{:02X}", byte)),
        }
    }

    encoded
}

fn virtual_url_for_file(path: &Path) -> Option<String> {
    let relative = path_relative_to_root(path)?;
    Some(format!(
        "https://mdview.example/{}",
        percent_encode_path(&relative)
    ))
}

fn content_type_for_path(path: &Path) -> &'static str {
    mime_guess::from_path(path)
        .first_raw()
        .unwrap_or("application/octet-stream")
}

fn create_web_resource_response(
    environment: &ICoreWebView2Environment,
    bytes: &[u8],
    status_code: i32,
    reason: &str,
    content_type: &str,
) -> Option<ICoreWebView2WebResourceResponse> {
    let stream = unsafe { SHCreateMemStream(Some(bytes)) }?;
    let reason_wide = pwstr_from_str(reason);
    let headers = format!(
        "Content-Type: {}\r\nAccess-Control-Allow-Origin: *",
        content_type
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
    )
}

fn create_virtual_resource_response(
    environment: &ICoreWebView2Environment,
    file_path: &Path,
    dark_mode: bool,
) -> Option<ICoreWebView2WebResourceResponse> {
    let ext = file_path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase())
        .unwrap_or_default();

    if ext == "md" || ext == "markdown" {
        let markdown_content = std::fs::read_to_string(file_path).ok()?;
        let html_body = crate::markdown::markdown_to_html(&markdown_content);
        let full_html = crate::markdown::wrap_html(&html_body, dark_mode);
        create_web_resource_response(
            environment,
            full_html.as_bytes(),
            200,
            "OK",
            "text/html; charset=utf-8",
        )
    } else {
        let bytes = std::fs::read(file_path).ok()?;
        create_web_resource_response(
            environment,
            &bytes,
            200,
            "OK",
            content_type_for_path(file_path),
        )
    }
}

fn init_webview2_sync(hwnd: HWND, html: &str) -> windows::core::Result<()> {
    log_debug(&format!("init_webview2_sync started, hwnd={:?}", hwnd.0));

    let html_owned = html.to_string();
    let controller_result: Rc<RefCell<Option<ICoreWebView2Controller>>> =
        Rc::new(RefCell::new(None));
    let error_result: Rc<RefCell<Option<windows::core::Error>>> = Rc::new(RefCell::new(None));
    let completed = Rc::new(RefCell::new(false));

    let controller_clone = controller_result.clone();
    let error_clone = error_result.clone();
    let completed_clone = completed.clone();

    // Create the environment first
    log_debug("Creating environment handler");
    let env_handler = CreateCoreWebView2EnvironmentCompletedHandler::create(Box::new(
        move |error_code, environment| {
            log_debug(&format!(
                "Environment callback fired, error_code={:?}",
                error_code
            ));
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
            let environment_for_resources = environment.clone();

            // Create the controller
            log_debug("Creating controller handler");
            let ctrl_handler = CreateCoreWebView2ControllerCompletedHandler::create(Box::new(
                move |error_code, controller| {
                    log_debug(&format!(
                        "Controller callback fired, error_code={:?}",
                        error_code
                    ));
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
                                let _ = settings.SetIsWebMessageEnabled(true); // needed for link clicks and ESC handling
                                let _ = settings.SetAreDefaultContextMenusEnabled(false); // disable right-click menu
                                let _ = settings.SetAreDevToolsEnabled(false); // disable F12 dev tools
                                let _ = settings.SetIsStatusBarEnabled(false); // disable status bar
                                let _ = settings.SetIsBuiltInErrorPageEnabled(false); // disable error pages
                                let _ = settings.SetAreDefaultScriptDialogsEnabled(false); // disable alert/confirm/prompt
                                // Zoom control left enabled for accessibility (Ctrl+scroll)
                            }

                            let filter = pwstr_from_str("https://mdview.example/*");
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
                                                    if let Some(response) =
                                                        create_response_for_virtual_url(
                                                            hwnd,
                                                            &resource_environment,
                                                            &uri,
                                                        )
                                                    {
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
                                            if !should_allow_webview_navigation(
                                                &uri,
                                                user_initiated.as_bool(),
                                            ) {
                                                let _ = args.SetCancel(true);
                                                open_url_in_browser(&uri);
                                            }
                                        }
                                    }
                                    Ok(())
                                }),
                            );
                            let mut nav_token: i64 = 0;
                            let _ = webview.add_NavigationStarting(&nav_handler, &mut nav_token);

                            if let Some(current_file) =
                                CURRENT_FILES.with(|f| f.borrow().get(&(hwnd.0 as isize)).cloned())
                            {
                                if let Some(url) = virtual_url_for_file(Path::new(&current_file)) {
                                    let url_wide = pwstr_from_str(&url);
                                    let _ = webview.Navigate(PCWSTR(url_wide.as_ptr()));
                                }
                            } else {
                                let html_wide = pwstr_from_str(&html_for_nav);
                                let _ = webview.NavigateToString(PCWSTR(html_wide.as_ptr()));
                            }

                            // Add message handler for link clicks and ESC to close
                            let viewer_hwnd = hwnd;
                            let handler = webview2_com::WebMessageReceivedEventHandler::create(
                                Box::new(move |_webview, args| {
                                    if let Some(args) = args {
                                        let mut message_ptr: windows::core::PWSTR =
                                            windows::core::PWSTR::null();
                                        if args.WebMessageAsJson(&mut message_ptr).is_ok()
                                            && !message_ptr.is_null()
                                        {
                                            let msg_str =
                                                message_ptr.to_string().unwrap_or_default();
                                            windows::Win32::System::Com::CoTaskMemFree(Some(
                                                message_ptr.0 as *const _,
                                            ));

                                            if msg_str.contains("\"close\"") {
                                                // Send WM_CLOSE to parent (TC lister window)
                                                if let Ok(parent) = GetParent(viewer_hwnd) {
                                                    if !parent.is_invalid() {
                                                        let _ = PostMessageW(
                                                            Some(parent),
                                                            WM_CLOSE,
                                                            WPARAM(0),
                                                            LPARAM(0),
                                                        );
                                                    }
                                                }
                                            } else if let Some(start) = msg_str.find("\"url\":\"") {
                                                let url_start = start + 7;
                                                if let Some(end) = msg_str[url_start..].find('"') {
                                                    let url = msg_str[url_start..url_start + end]
                                                        .replace("\\\\", "\\")
                                                        .replace("\\/", "/");

                                                    if msg_str.contains("openLink") {
                                                        open_resolved_or_external_url(
                                                            viewer_hwnd,
                                                            &url,
                                                        );
                                                    } else if msg_str.contains("followLink") {
                                                        handle_follow_link(viewer_hwnd, &url);
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
                                    let mut key_event_kind: COREWEBVIEW2_KEY_EVENT_KIND =
                                        COREWEBVIEW2_KEY_EVENT_KIND::default();

                                    if args.VirtualKey(&mut key).is_ok()
                                        && args.KeyEventKind(&mut key_event_kind).is_ok()
                                    {
                                        // Only handle key down events
                                        if key_event_kind == COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN
                                            || key_event_kind
                                                == COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN
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
                        let _ =
                            controller.add_AcceleratorKeyPressed(&accel_handler, &mut accel_token);
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
        let _ = CreateCoreWebView2EnvironmentWithOptions(
            None,
            PCWSTR(user_data_wide.as_ptr()),
            None,
            &env_handler,
        );
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
                log_debug(&format!(
                    "Message loop iteration {}, elapsed={:?}",
                    loop_count,
                    start.elapsed()
                ));
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
    log_debug(&format!(
        "Message loop exited, completed={}, loop_count={}",
        *completed.borrow(),
        loop_count
    ));

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

fn cleanup_viewer_state(hwnd: HWND) {
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow_mut().remove(&(hwnd.0 as isize)) {
            let _ = unsafe { controller.Close() };
        }
    });
    CURRENT_FILES.with(|f| {
        f.borrow_mut().remove(&(hwnd.0 as isize));
    });
    DARK_MODES.with(|m| {
        m.borrow_mut().remove(&(hwnd.0 as isize));
    });
}

fn unregister_window_class_if_unused() {
    let no_controllers = CONTROLLERS.with(|c| c.borrow().is_empty());
    let no_current_files = CURRENT_FILES.with(|f| f.borrow().is_empty());
    let no_dark_modes = DARK_MODES.with(|m| m.borrow().is_empty());

    if no_controllers && no_current_files && no_dark_modes {
        unsafe {
            if let Ok(hinstance) = GetModuleHandleW(None) {
                let class_name = window_class_name();
                let _ = UnregisterClassW(PCWSTR(class_name.as_ptr()), Some(hinstance.into()));
            }
        }
    }
}

pub fn close_window(hwnd: HWND) {
    // Remove and close the controller
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow_mut().remove(&(hwnd.0 as isize)) {
            let _ = unsafe { controller.Close() };
        }
    });

    // Destroy the window; final state cleanup happens in WM_NCDESTROY
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
        WM_DESTROY => LRESULT(0),
        WM_NCDESTROY => {
            cleanup_viewer_state(hwnd);
            unregister_window_class_if_unused();
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

fn hex_value(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn percent_decode(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut i = 0;

    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(hi), Some(lo)) = (hex_value(bytes[i + 1]), hex_value(bytes[i + 2])) {
                decoded.push((hi << 4) | lo);
                i += 3;
                continue;
            }
        }

        decoded.push(bytes[i]);
        i += 1;
    }

    String::from_utf8_lossy(&decoded).to_string()
}

fn percent_decode_repeated(input: &str) -> String {
    let mut current = input.to_string();

    for _ in 0..3 {
        let decoded = percent_decode(&current);
        if decoded == current {
            break;
        }
        current = decoded;
    }

    current
}

fn resolve_current_relative_path(hwnd: HWND, path: &str) -> Option<String> {
    CURRENT_FILES.with(|f| {
        let files = f.borrow();
        let current_path = files.get(&(hwnd.0 as isize))?;
        let base_dir = Path::new(current_path).parent()?;
        let target = base_dir.join(path);
        target
            .canonicalize()
            .ok()
            .map(|p| p.to_string_lossy().to_string())
    })
}

fn resolve_virtual_mdview_url(hwnd: HWND, url: &str) -> Option<String> {
    let relative_path = url
        .strip_prefix("https://mdview.example/")
        .or_else(|| url.strip_prefix("http://mdview.example/"))?
        .split(['?', '#'])
        .next()
        .unwrap_or_default();

    if relative_path.is_empty() {
        return None;
    }

    let decoded_path = percent_decode_repeated(relative_path);

    CURRENT_FILES.with(|f| {
        let files = f.borrow();
        let current_path = files.get(&(hwnd.0 as isize))?;
        let root_dir = get_path_root(Path::new(current_path))?;
        let target = root_dir.join(decoded_path);
        Some(target.to_string_lossy().to_string())
    })
}

fn create_response_for_virtual_url(
    hwnd: HWND,
    environment: &ICoreWebView2Environment,
    url: &str,
) -> Option<ICoreWebView2WebResourceResponse> {
    let resolved_path = resolve_virtual_mdview_url(hwnd, url)?;
    let file_path = Path::new(&resolved_path);
    if !file_path.exists() {
        return create_not_found_response(environment);
    }

    let dark_mode =
        DARK_MODES.with(|m| m.borrow().get(&(hwnd.0 as isize)).copied().unwrap_or(false));
    create_virtual_resource_response(environment, file_path, dark_mode)
}

fn is_internal_virtual_url(url: &str) -> bool {
    url.starts_with("https://mdview.example/") || url.starts_with("http://mdview.example/")
}

fn should_allow_webview_navigation(url: &str, user_initiated: bool) -> bool {
    !user_initiated
        || is_internal_virtual_url(url)
        || url == "about:blank"
        || url.starts_with("about:blank#")
        || url.starts_with("data:")
}

fn open_resolved_or_external_url(hwnd: HWND, url: &str) {
    if let Some(resolved_path) = resolve_virtual_mdview_url(hwnd, url) {
        open_url_in_browser(&resolved_path);
    } else {
        open_url_in_browser(url);
    }
}

fn load_file_into_viewer(hwnd: HWND, file_path: &str) {
    if std::fs::metadata(file_path).is_err() {
        return;
    }

    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow().get(&(hwnd.0 as isize)) {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                if let Some(url) = virtual_url_for_file(Path::new(file_path)) {
                    let url_wide = pwstr_from_str(&url);
                    let _ = unsafe { webview.Navigate(PCWSTR(url_wide.as_ptr())) };
                }
            }
        }
    });

    CURRENT_FILES.with(|f| {
        f.borrow_mut()
            .insert(hwnd.0 as isize, file_path.to_string());
    });
}

fn handle_follow_link(hwnd: HWND, url: &str) {
    if let Some(resolved_path) = resolve_virtual_mdview_url(hwnd, url) {
        let ext = Path::new(&resolved_path)
            .extension()
            .map(|e| e.to_string_lossy().to_lowercase())
            .unwrap_or_default();

        if ext == "md" || ext == "markdown" {
            load_file_into_viewer(hwnd, &resolved_path);
            return;
        }

        open_url_in_browser(&resolved_path);
        return;
    }

    if url.starts_with("http://") || url.starts_with("https://") || url.starts_with("mailto:") {
        open_url_in_browser(url);
        return;
    }

    let path_part = url.split('#').next().unwrap_or(url);
    let ext = Path::new(path_part)
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    if ext == "md" || ext == "markdown" {
        if let Some(resolved_path) = resolve_current_relative_path(hwnd, path_part) {
            load_file_into_viewer(hwnd, &resolved_path);
            return;
        }
    }

    if let Some(resolved_path) = resolve_current_relative_path(hwnd, path_part) {
        open_url_in_browser(&resolved_path);
        return;
    }

    open_url_in_browser(url);
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
