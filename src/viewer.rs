use std::cell::RefCell;
use std::collections::HashMap;
use std::env;
use std::rc::Rc;
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
    let html_owned = html.to_string();
    let controller_result: Rc<RefCell<Option<ICoreWebView2Controller>>> = Rc::new(RefCell::new(None));
    let error_result: Rc<RefCell<Option<windows::core::Error>>> = Rc::new(RefCell::new(None));
    let completed = Rc::new(RefCell::new(false));

    let controller_clone = controller_result.clone();
    let error_clone = error_result.clone();
    let completed_clone = completed.clone();

    // Create the environment first
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

            let controller_inner = controller_clone.clone();
            let error_inner = error_clone.clone();
            let completed_inner = completed_clone.clone();
            let html_for_nav = html_owned.clone();

            // Create the controller
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
                    }

                    *controller_inner.borrow_mut() = Some(controller);
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

    // Start the creation process - use TEMP folder (auto-cleaned by Windows)
    let user_data_folder = env::var("TEMP")
        .map(|p| format!("{}\\MDView_WebView2", p))
        .unwrap_or_else(|_| ".".to_string());
    let user_data_wide = pwstr_from_str(&user_data_folder);
    unsafe {
        let _ = CreateCoreWebView2EnvironmentWithOptions(None, PCWSTR(user_data_wide.as_ptr()), None, &env_handler);
    }

    // Pump messages until completion
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
