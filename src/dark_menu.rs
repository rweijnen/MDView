//! Dark mode menu bar support using undocumented UAH messages
//! Based on https://github.com/adzm/win32-custom-menubar-aero-theme

#![allow(dead_code)]

use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::*;
use windows::Win32::UI::WindowsAndMessaging::*;
use std::mem;

// Undocumented UAH messages for menu bar theming
pub const WM_UAHDRAWMENU: u32 = 0x0091;
pub const WM_UAHDRAWMENUITEM: u32 = 0x0092;
pub const WM_UAHMEASUREMENUITEM: u32 = 0x0094;

// Owner draw state flags (not always exposed in windows-rs)
const ODS_SELECTED: u32 = 0x0001;
const ODS_GRAYED: u32 = 0x0002;
const ODS_DISABLED: u32 = 0x0004;
const ODS_HOTLIGHT: u32 = 0x0040;

// UAH menu bar item metrics
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct UAHMENUITEMMETRICS {
    pub cx: [u32; 2],  // Size for menu bar items
}

// UAH menu popup metrics
#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct UAHMENUPOPUPMETRICS {
    pub rgcx: [u32; 4],  // Size for popup items
}

// UAH menu structure passed with WM_UAHDRAWMENU
#[repr(C)]
pub struct UAHMENU {
    pub hmenu: HMENU,
    pub hdc: HDC,
    pub dw_flags: u32,
}

// UAH menu item structure
#[repr(C)]
pub struct UAHMENUITEM {
    pub i_position: i32,  // 0-based position in menu bar
    pub umim: UAHMENUITEMMETRICS,
    pub umpm: UAHMENUPOPUPMETRICS,
}

// DRAWITEMSTRUCT for owner-drawn items
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DRAWITEMSTRUCT {
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
    pub dis: DRAWITEMSTRUCT,
    pub um: UAHMENU,
    pub umi: UAHMENUITEM,
}

/// Dark mode menu bar state
pub struct DarkMenuBar {
    enabled: bool,
}

impl DarkMenuBar {
    pub fn new() -> Self {
        Self {
            enabled: false,
        }
    }

    /// Enable dark mode menu bar for a window
    pub fn enable(&mut self, _hwnd: HWND) {
        self.enabled = true;
    }

    /// Disable and cleanup
    pub fn disable(&mut self) {
        self.enabled = false;
    }

    /// Check if enabled
    pub fn is_enabled(&self) -> bool {
        self.enabled
    }

    /// Handle UAH messages in window proc
    /// Returns Some(LRESULT) if message was handled, None otherwise
    pub fn handle_message(
        &self,
        hwnd: HWND,
        msg: u32,
        _wparam: WPARAM,
        lparam: LPARAM,
    ) -> Option<LRESULT> {
        if !self.enabled {
            return None;
        }

        match msg {
            WM_UAHDRAWMENU => {
                self.draw_menu_bar(lparam);
                Some(LRESULT(0))
            }
            WM_UAHDRAWMENUITEM => {
                self.draw_menu_item(hwnd, lparam);
                Some(LRESULT(0))
            }
            _ => None,
        }
    }

    /// Draw the menu bar background
    fn draw_menu_bar(&self, lparam: LPARAM) {
        unsafe {
            let uah_menu = &*(lparam.0 as *const UAHMENU);

            // Fill with dark background color (hardcoded dark gray)
            let mut rc = RECT::default();
            GetClipBox(uah_menu.hdc, &mut rc);

            // Use Windows 10/11 dark mode color: #202020
            let brush = CreateSolidBrush(COLORREF(0x00202020));
            FillRect(uah_menu.hdc, &rc, brush);
            let _ = DeleteObject(brush.into());
        }
    }

    /// Draw a single menu bar item
    fn draw_menu_item(&self, _hwnd: HWND, lparam: LPARAM) {
        unsafe {
            let uah_draw = &*(lparam.0 as *const UAHDRAWMENUITEM);
            let dis = &uah_draw.dis;
            let hdc = uah_draw.um.hdc;
            let mut rc = dis.rc_item;

            // Determine item state
            let item_state = dis.item_state;
            let is_hot = (item_state & ODS_HOTLIGHT) != 0;
            let is_selected = (item_state & ODS_SELECTED) != 0;
            let is_disabled = (item_state & (ODS_GRAYED | ODS_DISABLED)) != 0;

            // Draw background - dark gray normally, lighter when hot/selected
            let bg_color = if is_hot || is_selected {
                COLORREF(0x00404040) // Lighter gray for hover
            } else {
                COLORREF(0x00202020) // Dark gray background
            };
            let brush = CreateSolidBrush(bg_color);
            FillRect(hdc, &rc, brush);
            let _ = DeleteObject(brush.into());

            // Get menu item text
            let mut buffer = [0u16; 256];
            let mut mii = MENUITEMINFOW {
                cbSize: mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_STRING,
                dwTypeData: windows::core::PWSTR(buffer.as_mut_ptr()),
                cch: buffer.len() as u32,
                ..Default::default()
            };

            let hmenu = uah_draw.um.hmenu;
            let pos = uah_draw.umi.i_position as u32;

            if GetMenuItemInfoW(hmenu, pos, true, &mut mii).is_ok() {
                // Draw the text
                let text_color = if is_disabled {
                    COLORREF(0x00808080) // Gray
                } else {
                    COLORREF(0x00FFFFFF) // White
                };

                let old_bk_mode = SetBkMode(hdc, TRANSPARENT);
                let old_text_color = SetTextColor(hdc, text_color);

                // Adjust rect for text padding
                rc.left += 10;
                rc.right -= 10;

                // Draw text centered
                let text_len = buffer.iter().position(|&c| c == 0).unwrap_or(0);
                DrawTextW(
                    hdc,
                    &mut buffer[..text_len],
                    &mut rc,
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE,
                );

                SetTextColor(hdc, old_text_color);
                SetBkMode(hdc, BACKGROUND_MODE(old_bk_mode as u32));
            }
        }
    }
}

impl Drop for DarkMenuBar {
    fn drop(&mut self) {
        self.disable();
    }
}

impl Default for DarkMenuBar {
    fn default() -> Self {
        Self::new()
    }
}
