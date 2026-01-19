//! Dark mode menu bar support using undocumented UAH messages
//! Based on the menu_test program approach

use windows::core::PWSTR;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use std::mem;

// Undocumented UAH messages for menu bar theming
pub const WM_UAHDRAWMENU: u32 = 0x0091;
pub const WM_UAHDRAWMENUITEM: u32 = 0x0092;

// Owner draw state flags
const ODS_SELECTED: u32 = 0x0001;
const ODS_HOTLIGHT: u32 = 0x0040;

// Dark theme colors (must match main.rs)
const DARK_BG: u32 = 0x002D2D2D;
const DARK_HOVER: u32 = 0x00404040;

// UAH menu structure passed with WM_UAHDRAWMENU
#[repr(C)]
pub struct UAHMENU {
    pub hmenu: HMENU,
    pub hdc: HDC,
    pub dw_flags: u32,
}

// UAH menu item metrics
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct UAHMENUITEMMETRICS {
    pub cx: [u32; 2],
}

// UAH menu popup metrics
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct UAHMENUPOPUPMETRICS {
    pub rgcx: [u32; 4],
}

// UAH menu item structure
#[repr(C)]
pub struct UAHMENUITEM {
    pub i_position: i32,
    pub umim: UAHMENUITEMMETRICS,
    pub umpm: UAHMENUPOPUPMETRICS,
}

// DRAWITEMSTRUCT for owner-drawn items
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DRAWITEMSTRUCT_UAH {
    pub ctl_type: u32,
    pub ctl_id: u32,
    pub item_id: u32,
    pub item_action: u32,
    pub item_state: u32,
    pub hwnd_item: HWND,
    pub hdc: HDC,
    pub rc_item: RECT,
    pub item_data: usize,
}

// Combined structure for WM_UAHDRAWMENUITEM
#[repr(C)]
pub struct UAHDRAWMENUITEM {
    pub dis: DRAWITEMSTRUCT_UAH,
    pub um: UAHMENU,
    pub umi: UAHMENUITEM,
}

/// Paint over the menu bar separator line
pub fn paint_menu_separator(hwnd: HWND) {
    unsafe {
        let hdc = GetWindowDC(Some(hwnd));
        if !hdc.is_invalid() {
            let mut mbi = MENUBARINFO {
                cbSize: mem::size_of::<MENUBARINFO>() as u32,
                ..Default::default()
            };
            if GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mut mbi).is_ok() {
                let mut window_rect = RECT::default();
                let _ = GetWindowRect(hwnd, &mut window_rect);

                // Draw a dark line at the bottom of the menu bar
                let brush = CreateSolidBrush(COLORREF(DARK_BG));
                let line_rect = RECT {
                    left: 0,
                    top: mbi.rcBar.bottom - window_rect.top,
                    right: window_rect.right - window_rect.left,
                    bottom: mbi.rcBar.bottom - window_rect.top + 2,
                };
                FillRect(hdc, &line_rect, brush);
                let _ = DeleteObject(brush.into());
            }
            ReleaseDC(Some(hwnd), hdc);
        }
    }
}

/// Handle WM_UAHDRAWMENU - draw menu bar background
pub fn handle_uah_draw_menu(hwnd: HWND, lparam: LPARAM) -> LRESULT {
    unsafe {
        let pudm = lparam.0 as *const UAHMENU;
        if !pudm.is_null() {
            let udm = &*pudm;

            // Get menu bar rect using GetMenuBarInfo
            let mut mbi = MENUBARINFO {
                cbSize: mem::size_of::<MENUBARINFO>() as u32,
                ..Default::default()
            };
            if GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mut mbi).is_ok() {
                // Get window rect to convert to client coordinates
                let mut window_rect = RECT::default();
                let _ = GetWindowRect(hwnd, &mut window_rect);

                // Convert menu bar rect to window-relative coordinates
                // Extend by a few pixels to cover any separator lines
                let rc = RECT {
                    left: mbi.rcBar.left - window_rect.left,
                    top: mbi.rcBar.top - window_rect.top,
                    right: mbi.rcBar.right - window_rect.left,
                    bottom: mbi.rcBar.bottom - window_rect.top + 4,
                };

                // Fill menu bar background with dark color
                let brush = CreateSolidBrush(COLORREF(DARK_BG));
                FillRect(udm.hdc, &rc, brush);
                let _ = DeleteObject(brush.into());
            }
        }
    }
    LRESULT(0)
}

/// Handle WM_UAHDRAWMENUITEM - draw menu bar item
pub fn handle_uah_draw_menu_item(lparam: LPARAM) -> LRESULT {
    unsafe {
        let pudmi = lparam.0 as *mut UAHDRAWMENUITEM;
        if !pudmi.is_null() {
            let udmi = &mut *pudmi;
            let dis = &udmi.dis;

            // Choose color based on state
            let bg_color = if (dis.item_state & ODS_SELECTED) != 0 ||
                              (dis.item_state & ODS_HOTLIGHT) != 0 {
                DARK_HOVER
            } else {
                DARK_BG
            };

            // Fill background
            let brush = CreateSolidBrush(COLORREF(bg_color));
            FillRect(dis.hdc, &dis.rc_item, brush);
            let _ = DeleteObject(brush.into());

            // Get menu item text
            let mut text_buf = [0u16; 256];
            let mut mii = MENUITEMINFOW {
                cbSize: mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_STRING,
                dwTypeData: PWSTR(text_buf.as_mut_ptr()),
                cch: 256,
                ..Default::default()
            };
            if GetMenuItemInfoW(udmi.um.hmenu, udmi.umi.i_position as u32, true, &mut mii).is_ok() {
                // Get system menu font and scale up slightly to match popup menu size
                let menu_font = {
                    let mut ncm = NONCLIENTMETRICSW {
                        cbSize: mem::size_of::<NONCLIENTMETRICSW>() as u32,
                        ..Default::default()
                    };
                    if SystemParametersInfoW(
                        SPI_GETNONCLIENTMETRICS,
                        ncm.cbSize,
                        Some(&mut ncm as *mut _ as *mut _),
                        SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
                    ).is_ok() {
                        // Scale up font height to match popup menu (lfHeight is negative)
                        let mut lf = ncm.lfMenuFont;
                        lf.lfHeight = (lf.lfHeight as f32 * 1.12) as i32;
                        CreateFontIndirectW(&lf)
                    } else {
                        // Fallback to hardcoded font if system call fails
                        let font_name: Vec<u16> = "Segoe UI\0".encode_utf16().collect();
                        CreateFontW(
                            -16, 0, 0, 0,
                            FW_NORMAL.0 as i32, 0, 0, 0,
                            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                            CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                            DEFAULT_PITCH.0 as u32,
                            windows::core::PCWSTR(font_name.as_ptr()),
                        )
                    }
                };
                let old_font = SelectObject(dis.hdc, menu_font.into());

                let old_bk = SetBkMode(dis.hdc, TRANSPARENT);
                let old_color = SetTextColor(dis.hdc, COLORREF(0x00FFFFFF));

                let mut rc = dis.rc_item;
                let format = DT_CENTER | DT_VCENTER | DT_SINGLELINE;
                DrawTextW(dis.hdc, &mut text_buf[..mii.cch as usize], &mut rc, format);

                SetTextColor(dis.hdc, old_color);
                SetBkMode(dis.hdc, BACKGROUND_MODE(old_bk as u32));
                SelectObject(dis.hdc, old_font);
                let _ = DeleteObject(menu_font.into());
            }
        }
    }
    LRESULT(0)
}
