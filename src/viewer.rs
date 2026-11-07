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
pub(crate) fn log_debug(msg: &str) {
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
use windows::Win32::UI::Shell::SHCreateMemStream;
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{Interface, PCWSTR};

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
    /// Links waiting to be handled outside the WebView2 event that reported them.
    static PENDING_LINKS: RefCell<std::collections::VecDeque<(isize, String)>> = RefCell::new(std::collections::VecDeque::new());
    /// Viewers that are about to call NavigateToString, so only that
    /// navigation to a data:/about:blank page is allowed.
    static EXPECT_STRING_PAGE: RefCell<std::collections::HashSet<isize>> = RefCell::new(std::collections::HashSet::new());
    /// Per viewer: the document for which the user chose "Load images".
    static REMOTE_IMAGES_ONCE: RefCell<HashMap<isize, std::path::PathBuf>> = RefCell::new(HashMap::new());
    /// Per viewer: whether the page being shown may load remote images (set when MDView serves it).
    static PAGE_ALLOWS_REMOTE_IMAGES: RefCell<HashMap<isize, bool>> = RefCell::new(HashMap::new());
    /// Per viewer: the bars MDView put on the page being shown; only answers to these are accepted.
    static PAGE_BARS: RefCell<HashMap<isize, Vec<crate::markdown::Bar>>> = RefCell::new(HashMap::new());
    /// Per viewer: what the user chose while the document being shown is open (#24).
    static DOCUMENT_CHOICES: RefCell<HashMap<isize, crate::markdown::DocumentChoices>> = RefCell::new(HashMap::new());
}

/// Posted to a viewer window to handle its queued links (see `queue_link`).
const WM_MDVIEW_LINK: u32 = WM_APP + 1;
/// A file from a network location was blocked: tell the page.
const WM_MDVIEW_NETWORK: u32 = WM_APP + 2;

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

/// The URL a document is shown at (see `vpath`).
fn virtual_url_for_file(path: &Path) -> Option<String> {
    crate::vpath::document_url(path)
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

/// The mark of a downloaded document that MDView must ask about before
/// showing it in this viewer: not if the user chose "Don't ask again" or
/// already chose Open while this document is open.
fn download_gate(hwnd: HWND, file_path: &Path) -> Option<crate::zone::Mark> {
    if crate::settings::open_downloaded_without_asking() {
        return None;
    }
    let mark = crate::zone::mark_of_the_web(file_path)?;
    let opened = DOCUMENT_CHOICES
        .with(|c| c.borrow_mut().entry(hwnd.0 as isize).or_default().opened_download(file_path));
    (!opened).then_some(mark)
}

fn create_virtual_resource_response(
    hwnd: HWND,
    environment: &ICoreWebView2Environment,
    file_path: &Path,
    dark_mode: bool,
) -> Option<ICoreWebView2WebResourceResponse> {
    if crate::links::is_markdown(file_path) {
        let (html_body, remote_images, bars) = if let Some(mark) = download_gate(hwnd, file_path) {
            // Downloaded from the internet (#24): nothing from the document
            // is shown until the user chooses to open it.
            let gate = crate::markdown::download_gate_html(&file_path.display().to_string(), mark.host_url.as_deref());
            (gate, crate::markdown::RemoteImages::Blocked, vec![crate::markdown::download_gate_bar()])
        } else {
            match crate::files::read_text(file_path) {
                Ok(markdown_content) => {
                    let remote_images = remote_images_for(hwnd, file_path);
                    // The plugin never asks to register (#37): TC is the host.
                    let bars = DOCUMENT_CHOICES.with(|c| {
                        let mut all = c.borrow_mut();
                        let choices = all.entry(hwnd.0 as isize).or_default();
                        choices.page_served(file_path);
                        let ask_network_files = crate::settings::network_files_setting()
                            == crate::settings::AskSetting::Ask
                            && !choices.network_loaded(file_path);
                        crate::markdown::document_bars(&crate::markdown::PageBars {
                            remote_images,
                            ask_to_register: false,
                            ask_network_files,
                            closed: choices.closed_bars(file_path),
                        })
                    });
                    (crate::markdown::markdown_to_page_html(&markdown_content, file_path), remote_images, bars)
                }
                Err(reason) => (
                    crate::markdown::read_error_html(&file_path.display().to_string(), &reason),
                    crate::markdown::RemoteImages::Blocked,
                    Vec::new(),
                ),
            }
        };
        PAGE_ALLOWS_REMOTE_IMAGES.with(|a| {
            a.borrow_mut().insert(
                hwnd.0 as isize,
                remote_images == crate::markdown::RemoteImages::Allowed,
            )
        });
        let full_html = crate::markdown::wrap_html(&html_body, dark_mode, remote_images, &bars);
        PAGE_BARS.with(|b| b.borrow_mut().insert(hwnd.0 as isize, bars));
        create_web_resource_response(
            environment,
            full_html.as_bytes(),
            200,
            "OK",
            "text/html; charset=utf-8",
            "",
        )
    } else {
        let bytes = match crate::files::read_bytes(file_path) {
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
        concat!("Blocked by ", crate::settings::app_name!()),
        "text/plain; charset=utf-8",
        PASSIVE_FILE_HEADERS,
    )
}

/// A request outside MDView's own files may only go out if it is an image on
/// a page whose remote images the user chose to load.
fn remote_request_allowed(hwnd: HWND, uri: &str, context: COREWEBVIEW2_WEB_RESOURCE_CONTEXT) -> bool {
    let lower = uri.to_ascii_lowercase();
    context == COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE
        && (lower.starts_with("https://") || lower.starts_with("http://"))
        && PAGE_ALLOWS_REMOTE_IMAGES
            .with(|a| a.borrow().get(&(hwnd.0 as isize)).copied().unwrap_or(false))
}

/// Remote images for a document in this viewer: the saved choice, or "load
/// once" if the user chose that for this document in the banner.
fn remote_images_for(hwnd: HWND, file_path: &Path) -> crate::markdown::RemoteImages {
    use crate::markdown::RemoteImages;
    use crate::settings::AskSetting;
    match crate::settings::remote_images_setting() {
        AskSetting::Always => RemoteImages::Allowed,
        AskSetting::Never => RemoteImages::Blocked,
        AskSetting::Ask => {
            let once = REMOTE_IMAGES_ONCE
                .with(|o| o.borrow().get(&(hwnd.0 as isize)).is_some_and(|p| p == file_path));
            if once { RemoteImages::Allowed } else { RemoteImages::AskToLoad }
        }
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
                                let _ = settings.SetAreHostObjectsAllowed(false);
                                let _ = settings.SetIsBuiltInErrorPageEnabled(false); // disable error pages
                                let _ = settings.SetAreDefaultScriptDialogsEnabled(false); // disable alert/confirm/prompt
                                // Zoom control left enabled for accessibility (Ctrl+scroll)
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
                                                            hwnd,
                                                            &resource_environment,
                                                            &uri,
                                                        )
                                                    } else {
                                                        let mut context =
                                                            COREWEBVIEW2_WEB_RESOURCE_CONTEXT::default();
                                                        let _ = args.ResourceContext(&mut context);
                                                        if remote_request_allowed(hwnd, &uri, context) {
                                                            None
                                                        } else {
                                                            log_debug(&format!("blocked request to {uri}"));
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
                                            if !navigation_allowed(hwnd, &uri) {
                                                let _ = args.SetCancel(true);
                                                log_debug(&format!("blocked navigation to {uri}"));
                                                if user_initiated.as_bool() {
                                                    queue_link(hwnd, uri);
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
                                                log_debug(&format!("blocked new window for {uri}"));
                                                if user_initiated.as_bool() {
                                                    queue_link(hwnd, uri);
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
                                                log_debug("blocked download");
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

                            if let Some(current_file) =
                                CURRENT_FILES.with(|f| f.borrow().get(&(hwnd.0 as isize)).cloned())
                            {
                                if let Some(url) = virtual_url_for_file(Path::new(&current_file)) {
                                    let url_wide = pwstr_from_str(&url);
                                    let _ = webview.Navigate(PCWSTR(url_wide.as_ptr()));
                                }
                            } else {
                                let html_wide = pwstr_from_str(&html_for_nav);
                                EXPECT_STRING_PAGE
                                    .with(|e| e.borrow_mut().insert(hwnd.0 as isize));
                                let _ = webview.NavigateToString(PCWSTR(html_wide.as_ptr()));
                            }

                            // Messages from MDView's own page script only (links, ESC to close).
                            let viewer_hwnd = hwnd;
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
                                        if !is_trusted_page(viewer_hwnd, &source) {
                                            log_debug(&format!("ignored message from {source}"));
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
                                            if let Some(msg) = crate::links::parse_page_message(&json) {
                                                handle_page_message(viewer_hwnd, &source, &msg);
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

    // Own WebView2 data folder for this host program (#33)
    let user_data_folder = match crate::settings::webview2_data_folder() {
        Ok(folder) => folder.to_string_lossy().into_owned(),
        Err(e) => {
            log_debug(&format!("no WebView2 data folder: {e}"));
            return Err(windows::core::Error::new(E_FAIL, e));
        }
    };
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
    let key = hwnd.0 as isize;
    PENDING_LINKS.with(|q| q.borrow_mut().retain(|(k, _)| *k != key));
    EXPECT_STRING_PAGE.with(|e| e.borrow_mut().remove(&key));
    REMOTE_IMAGES_ONCE.with(|o| o.borrow_mut().remove(&key));
    PAGE_ALLOWS_REMOTE_IMAGES.with(|a| a.borrow_mut().remove(&key));
    PAGE_BARS.with(|b| b.borrow_mut().remove(&key));
    DOCUMENT_CHOICES.with(|c| c.borrow_mut().remove(&key));
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
        WM_MDVIEW_LINK => {
            handle_queued_links(hwnd);
            LRESULT(0)
        }
        WM_MDVIEW_NETWORK => {
            post_network_blocked(hwnd);
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

/// The file a URL from the page stands for while this viewer's file is
/// shown, or None if it does not name a plain file (see `vpath`).
fn resolve_virtual_mdview_url(hwnd: HWND, url: &str) -> Option<String> {
    CURRENT_FILES.with(|f| {
        let files = f.borrow();
        let current = files.get(&(hwnd.0 as isize))?;
        crate::vpath::resolve(url, Path::new(current)).map(|path| path.to_string_lossy().into_owned())
    })
}

fn create_response_for_virtual_url(
    hwnd: HWND,
    environment: &ICoreWebView2Environment,
    url: &str,
) -> Option<ICoreWebView2WebResourceResponse> {
    // A URL that names no plain file is answered here, never sent anywhere.
    let Some(resolved_path) = resolve_virtual_mdview_url(hwnd, url) else {
        return create_not_found_response(environment);
    };
    let file_path = Path::new(&resolved_path);
    // Decided before the file is touched: MDView does not connect to a
    // network location until the user agrees.
    if !network_file_allowed(hwnd, file_path) {
        return create_blocked_response(environment);
    }
    if !file_path.exists() {
        return create_not_found_response(environment);
    }

    let dark_mode =
        DARK_MODES.with(|m| m.borrow().get(&(hwnd.0 as isize)).copied().unwrap_or(false));
    create_virtual_resource_response(hwnd, environment, file_path, dark_mode)
}

/// Whether a file this viewer's page asks for may be read. Files on another
/// share or a network drive wait for the user's answer in the network bar.
fn network_file_allowed(hwnd: HWND, path: &Path) -> bool {
    use crate::settings::AskSetting;
    use crate::vpath::Location;
    let key = hwnd.0 as isize;
    let Some(current) = CURRENT_FILES.with(|f| f.borrow().get(&key).cloned()) else { return false };
    let location = match crate::vpath::location(path, Path::new(&current), crate::vpath::drive_kind) {
        Some(Location::Network(location)) => location,
        Some(Location::SameRoot | Location::Local) => return true,
        None => return false,
    };
    match crate::settings::network_files_setting() {
        AskSetting::Always => true,
        AskSetting::Never => false,
        AskSetting::Ask => DOCUMENT_CHOICES.with(|c| {
            let mut all = c.borrow_mut();
            let choices = all.entry(key).or_default();
            let Some(document) = choices.document().map(Path::to_path_buf) else { return false };
            if choices.network_loaded(&document) {
                return true;
            }
            if choices.note_network_blocked(&document, &location) {
                log_debug(&format!("blocked a file from network location {location}"));
                // Told to the page from the window, not from inside this
                // WebView2 callback.
                let _ = unsafe { PostMessageW(Some(hwnd), WM_MDVIEW_NETWORK, WPARAM(0), LPARAM(0)) };
            }
            false
        }),
    }
}

/// The network location (another share or a network drive) `path` is on,
/// seen from this viewer's document; None for local files.
fn network_location(hwnd: HWND, path: &Path) -> Option<String> {
    let current = CURRENT_FILES.with(|f| f.borrow().get(&(hwnd.0 as isize)).cloned())?;
    match crate::vpath::location(path, Path::new(&current), crate::vpath::drive_kind)? {
        crate::vpath::Location::Network(location) => Some(location),
        crate::vpath::Location::SameRoot | crate::vpath::Location::Local => None,
    }
}
/// Tell this viewer's page which network locations were blocked, for its
/// network bar. The page ignores locations it already lists.
fn post_network_blocked(hwnd: HWND) {
    let key = hwnd.0 as isize;
    let locations = DOCUMENT_CHOICES.with(|c| {
        let mut all = c.borrow_mut();
        let choices = all.entry(key).or_default();
        match choices.document().map(Path::to_path_buf) {
            Some(document) => choices.network_blocked(&document),
            None => Vec::new(),
        }
    });
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow().get(&key) {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                for location in locations {
                    let json: Vec<u16> = crate::markdown::network_blocked_message(&location)
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
    use crate::settings::AskSetting;
    let save = |setting| match crate::settings::set_network_files_setting(setting) {
        Ok(()) => true,
        Err(e) => {
            show_error(hwnd, &format!("Could not save the network files setting:\n\n{e}"));
            false
        }
    };
    match choice {
        "once" => {
            DOCUMENT_CHOICES.with(|c| c.borrow_mut().entry(hwnd.0 as isize).or_default().load_network(document));
            reload_page(hwnd);
        }
        "always" => {
            if save(AskSetting::Always) {
                reload_page(hwnd);
            }
        }
        "never" => {
            save(AskSetting::Never);
        }
        _ => {}
    }
}

/// Any of MDView's own URLs: served by the host, never fetched from the web.
fn is_internal_virtual_url(url: &str) -> bool {
    crate::vpath::is_virtual_url(url)
}

/// Only MDView's own pages may be shown: Markdown documents on the virtual
/// origin, and a page MDView itself loads with NavigateToString. Everything
/// else (script or <meta refresh> navigation, other files, other sites) is
/// cancelled; a link the user clicked goes through `queue_link`.
fn navigation_allowed(hwnd: HWND, url: &str) -> bool {
    if crate::vpath::is_document_origin(url) {
        return resolve_virtual_mdview_url(hwnd, url)
            .is_some_and(|path| crate::links::is_markdown(Path::new(&path)));
    }
    if url.starts_with("data:") || url == "about:blank" {
        return EXPECT_STRING_PAGE.with(|e| e.borrow_mut().remove(&(hwnd.0 as isize)));
    }
    false
}

/// Whether a message comes from a page MDView itself loaded.
fn is_trusted_page(hwnd: HWND, source: &str) -> bool {
    if crate::vpath::is_document_origin(source) {
        return resolve_virtual_mdview_url(hwnd, source)
            .is_some_and(|path| crate::links::is_markdown(Path::new(&path)));
    }
    source.starts_with("data:") || source == "about:blank"
}

/// Handle a link outside of the WebView2 event that reported it, so the
/// confirmation dialog is not shown from inside a WebView2 callback.
fn queue_link(hwnd: HWND, url: String) {
    PENDING_LINKS.with(|q| q.borrow_mut().push_back((hwnd.0 as isize, url)));
    let _ = unsafe { PostMessageW(Some(hwnd), WM_MDVIEW_LINK, WPARAM(0), LPARAM(0)) };
}

fn handle_queued_links(hwnd: HWND) {
    let key = hwnd.0 as isize;
    loop {
        let next = PENDING_LINKS.with(|q| {
            let mut q = q.borrow_mut();
            let pos = q.iter().position(|(k, _)| *k == key)?;
            q.remove(pos).map(|(_, url)| url)
        });
        let Some(url) = next else { break };
        // Dialogs belong to Total Commander's Lister window.
        let owner = unsafe { GetParent(hwnd) }.ok().filter(|p| !p.is_invalid()).unwrap_or(hwnd);
        let action = crate::links::classify_url(&url, |u| {
            resolve_virtual_mdview_url(hwnd, u).map(std::path::PathBuf::from)
        }, |p| network_location(hwnd, p));
        crate::links::perform(action, owner, |path| {
            load_file_into_viewer(hwnd, &path.to_string_lossy())
        });
    }
}

/// A message from MDView's page script.
fn handle_page_message(hwnd: HWND, source: &str, msg: &HashMap<String, String>) {
    match msg.get("type").map(String::as_str) {
        Some("close") => close_lister(hwnd),
        Some("openLink") => {
            if let Some(url) = msg.get("url") {
                queue_link(hwnd, url.clone());
            }
        }
        Some("networkStatus") => post_network_blocked(hwnd),
        Some("bar") => {
            let (Some(bar), Some(choice)) = (msg.get("bar"), msg.get("choice")) else { return };
            let offered = PAGE_BARS.with(|b| {
                b.borrow()
                    .get(&(hwnd.0 as isize))
                    .is_some_and(|bars| crate::markdown::bar_choice_offered(bars, bar, choice))
            });
            if !offered {
                log_debug(&format!("ignored answer to a bar not on this page: {bar} / {choice}"));
                return;
            }
            if choice == crate::markdown::CLOSE_CHOICE {
                if let Some(path) = resolve_virtual_mdview_url(hwnd, source) {
                    DOCUMENT_CHOICES.with(|c| {
                        c.borrow_mut().entry(hwnd.0 as isize).or_default().close_bar(Path::new(&path), bar)
                    });
                }
                return;
            }
            match bar.as_str() {
                crate::markdown::DOWNLOAD_GATE => {
                    if let Some(path) = resolve_virtual_mdview_url(hwnd, source) {
                        handle_download_gate_choice(hwnd, choice, Path::new(&path));
                    }
                }
                crate::markdown::REMOTE_IMAGES_BAR => {
                    if let Some(path) = resolve_virtual_mdview_url(hwnd, source) {
                        handle_remote_images_choice(hwnd, choice, Path::new(&path));
                    }
                }
                crate::markdown::NETWORK_FILES_BAR => {
                    if let Some(path) = resolve_virtual_mdview_url(hwnd, source) {
                        handle_network_files_choice(hwnd, choice, Path::new(&path));
                    }
                }
                _ => {}
            }
        }
        _ => {}
    }
}

/// Close Total Commander's Lister window.
fn close_lister(hwnd: HWND) {
    if let Ok(parent) = unsafe { GetParent(hwnd) } {
        if !parent.is_invalid() {
            let _ = unsafe { PostMessageW(Some(parent), WM_CLOSE, WPARAM(0), LPARAM(0)) };
        }
    }
}

/// The user's answer on the page shown instead of a downloaded document
/// (#24). The document is shown only if the choice worked.
fn handle_download_gate_choice(hwnd: HWND, choice: &str, document: &Path) {
    let result = match choice {
        "open" => {
            DOCUMENT_CHOICES
                .with(|c| c.borrow_mut().entry(hwnd.0 as isize).or_default().open_download(document));
            Ok(())
        }
        "unblock" => crate::zone::remove_mark(document)
            .map_err(|e| format!("Could not remove the mark of the web from\n{e}")),
        "never" => crate::settings::set_open_downloaded_without_asking()
            .map_err(|e| format!("Could not save \"Don't ask again\":\n\n{e}")),
        "closeWindow" => {
            close_lister(hwnd);
            return;
        }
        _ => return,
    };
    match result {
        Ok(()) => reload_page(hwnd),
        Err(message) => show_error(hwnd, &message),
    }
}

fn reload_page(hwnd: HWND) {
    CONTROLLERS.with(|c| {
        if let Some(controller) = c.borrow().get(&(hwnd.0 as isize)) {
            if let Ok(webview) = unsafe { controller.CoreWebView2() } {
                let _ = unsafe { webview.Reload() };
            }
        }
    });
}

/// The user's answer in the remote-images banner.
fn handle_remote_images_choice(hwnd: HWND, choice: &str, document: &Path) {
    use crate::settings::AskSetting;
    let reload = match choice {
        "once" => {
            REMOTE_IMAGES_ONCE
                .with(|o| o.borrow_mut().insert(hwnd.0 as isize, document.to_path_buf()));
            true
        }
        "always" => save_remote_images_setting(hwnd, AskSetting::Always),
        "never" => {
            save_remote_images_setting(hwnd, AskSetting::Never);
            false
        }
        _ => false,
    };
    if reload {
        reload_page(hwnd);
    }
}

fn save_remote_images_setting(hwnd: HWND, setting: crate::settings::AskSetting) -> bool {
    match crate::settings::set_remote_images_setting(setting) {
        Ok(()) => true,
        Err(e) => {
            show_error(hwnd, &format!("Could not save the remote images setting:\n\n{e}"));
            false
        }
    }
}

/// Something the user asked for failed: say so, owned by TC's Lister window.
fn show_error(hwnd: HWND, message: &str) {
    log_debug(message);
    let owner = unsafe { GetParent(hwnd) }.ok().filter(|p| !p.is_invalid()).unwrap_or(hwnd);
    let text: Vec<u16> = format!("{message}\0").encode_utf16().collect();
    let title: Vec<u16> = format!("{}\0", crate::settings::APP_NAME).encode_utf16().collect();
    unsafe {
        MessageBoxW(Some(owner), PCWSTR(text.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
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

#[cfg(test)]
mod tests {
    use super::*;

    // A fake window handle: the URL mapping only uses it as a lookup key.
    fn viewer_showing(file: &str) -> HWND {
        let hwnd = HWND(0x4D44 as *mut _);
        CURRENT_FILES.with(|f| f.borrow_mut().insert(hwnd.0 as isize, file.to_string()));
        hwnd
    }

    // Compared as text: the path is shown to the user (link questions, the
    // window title, recent files), so it must use `\` only (#49).
    fn assert_resolves(hwnd: HWND, url: &str, expected: &str) {
        let resolved = resolve_virtual_mdview_url(hwnd, url)
            .unwrap_or_else(|| panic!("{url} did not resolve to a file"));
        assert_eq!(resolved, expected, "resolved {url}");
    }

    fn assert_refused(hwnd: HWND, url: &str) {
        let resolved = resolve_virtual_mdview_url(hwnd, url);
        assert_eq!(resolved, None, "{url} must not resolve to a file");
    }

    // A URL from a document may only name a file below the root it maps to.
    // Encoded separators, drive letters, `..`, devices, pipes and alternate
    // data streams must never become a path elsewhere.
    #[test]
    fn test_urls_cannot_escape_to_other_paths() {
        let hwnd = viewer_showing(r"C:\docs\note.md");
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
            assert_refused(hwnd, url);
        }
        // Decoded once only: a double-encoded `\` is a literal name on the
        // document's own drive, never a separator.
        assert_resolves(hwnd, "https://mdview.example/%255C%255Cserver%255Cshare%255Cx.png", r"C:\%5C%5Cserver%5Cshare%5Cx.png");
    }

    #[test]
    fn test_document_url_is_path_from_drive_root() {
        assert_eq!(
            virtual_url_for_file(Path::new(r"C:\docs\note.md")).as_deref(),
            Some("https://mdview.example/docs/note.md")
        );
    }

    // Issue #7: `![test](img/screen.png)` next to the document. The browser
    // resolves the relative src against the document URL
    // (https://mdview.example/docs/note.md) and requests the URL below, which
    // must map back to the image file next to the document.
    #[test]
    fn test_relative_image_maps_to_file_next_to_document_issue_7() {
        let hwnd = viewer_showing(r"C:\docs\note.md");
        assert_resolves(
            hwnd,
            "https://mdview.example/docs/img/screen.png",
            r"C:\docs\img\screen.png",
        );
    }

    // Folder and file names with spaces and non-ASCII letters arrive
    // percent-encoded (UTF-8) from the browser.
    #[test]
    fn test_encoded_non_ascii_and_spaces_round_trip() {
        let doc = r"C:\Dokumenty\poznámky\note.md";
        assert_eq!(
            virtual_url_for_file(Path::new(doc)).as_deref(),
            Some("https://mdview.example/Dokumenty/pozn%C3%A1mky/note.md")
        );

        let hwnd = viewer_showing(doc);
        assert_resolves(
            hwnd,
            "https://mdview.example/Dokumenty/pozn%C3%A1mky/obr%C3%A1zky/sn%C3%ADmek%20obrazovky.png",
            r"C:\Dokumenty\poznámky\obrázky\snímek obrazovky.png",
        );
    }

    #[test]
    fn test_query_and_fragment_are_ignored() {
        let hwnd = viewer_showing(r"C:\docs\note.md");
        assert_resolves(
            hwnd,
            "https://mdview.example/docs/img/a.png?v=2#x",
            r"C:\docs\img\a.png",
        );
    }

    // Documents on a network share resolve images on the same share.
    #[test]
    fn test_unc_share_images_stay_on_share() {
        let doc = r"\\server\share\docs\note.md";
        assert_eq!(
            virtual_url_for_file(Path::new(doc)).as_deref(),
            Some("https://mdview.example/docs/note.md")
        );

        let hwnd = viewer_showing(doc);
        assert_resolves(
            hwnd,
            "https://mdview.example/docs/img/screen.png",
            r"\\server\share\docs\img\screen.png",
        );
    }
}
