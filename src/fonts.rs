//! The fonts installed on this PC, for the settings page (#50). Shared by the
//! standalone viewer and the TC plugin.

use windows::Win32::Foundation::LPARAM;
use windows::Win32::Graphics::Gdi::{
    DEFAULT_CHARSET, EnumFontFamiliesExW, GetDC, LOGFONTW, ReleaseDC, TEXTMETRICW,
};

/// An installed font family.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Font {
    pub name: String,
    /// Every character has the same width (for the code font list).
    pub monospace: bool,
}

/// The installed font families, sorted by name, each once.
pub fn installed() -> Vec<Font> {
    let mut found: Vec<Font> = Vec::new();
    unsafe {
        let hdc = GetDC(None);
        if hdc.is_invalid() {
            return Vec::new();
        }
        let filter = LOGFONTW { lfCharSet: DEFAULT_CHARSET, ..Default::default() };
        EnumFontFamiliesExW(hdc, &filter, Some(collect), LPARAM(&mut found as *mut Vec<Font> as isize), 0);
        ReleaseDC(None, hdc);
    }
    clean(found)
}

unsafe extern "system" fn collect(font: *const LOGFONTW, _metrics: *const TEXTMETRICW, _kind: u32, list: LPARAM) -> i32 {
    unsafe {
        let font = &*font;
        let list = &mut *(list.0 as *mut Vec<Font>);
        let len = font.lfFaceName.iter().position(|&c| c == 0).unwrap_or(font.lfFaceName.len());
        list.push(Font {
            name: String::from_utf16_lossy(&font.lfFaceName[..len]),
            // FIXED_PITCH in the low bits of lfPitchAndFamily.
            monospace: font.lfPitchAndFamily & 0x3 == 1,
        });
    }
    1
}

/// Sorted, each family once, without the vertical (`@`) variants Windows lists
/// for East Asian fonts and without names the page could not use.
pub fn clean(mut fonts: Vec<Font>) -> Vec<Font> {
    fonts.retain(|f| !f.name.starts_with('@') && crate::settings::is_font_name(&f.name));
    fonts.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()).then(a.name.cmp(&b.name)));
    fonts.dedup_by(|a, b| {
        if a.name.eq_ignore_ascii_case(&b.name) {
            b.monospace |= a.monospace;
            true
        } else {
            false
        }
    });
    fonts
}

#[cfg(test)]
mod tests {
    use super::*;

    fn font(name: &str, monospace: bool) -> Font {
        Font { name: name.to_string(), monospace }
    }

    #[test]
    fn test_font_list_is_sorted_once_without_vertical_variants() {
        let cleaned = clean(vec![
            font("Segoe UI", false),
            font("@MS Gothic", false),
            font("Consolas", true),
            font("arial", false),
            font("Arial", false),
            font("MS Gothic", true),
            font("Bad\"Name", false),
            font("Consolas", true),
        ]);
        let names: Vec<&str> = cleaned.iter().map(|f| f.name.as_str()).collect();
        assert_eq!(names, ["Arial", "Consolas", "MS Gothic", "Segoe UI"]);
        assert!(cleaned.iter().find(|f| f.name == "Consolas").unwrap().monospace);
    }

    // On any Windows PC: the system's own fonts are found.
    #[test]
    fn test_installed_fonts_include_windows_fonts() {
        let fonts = installed();
        assert!(fonts.iter().any(|f| f.name == "Segoe UI"), "{} fonts", fonts.len());
        assert!(fonts.iter().any(|f| f.name == "Consolas" && f.monospace));
    }
}
