//! Put text on the Windows clipboard ("Copy link address", #59).

use windows::Win32::Foundation::{GlobalFree, HANDLE, HGLOBAL, HWND};
use windows::Win32::System::DataExchange::{CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData};
use windows::Win32::System::Memory::{GMEM_MOVEABLE, GlobalAlloc, GlobalLock, GlobalUnlock};

const CF_UNICODETEXT: u32 = 13;

/// Replace the clipboard's contents with `text`. The error says why it failed
/// (for example another program keeps the clipboard open).
pub fn set_text(owner: HWND, text: &str) -> Result<(), String> {
    let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        OpenClipboard(Some(owner)).map_err(|e| format!("the clipboard is in use: {}", e.message()))?;
        let result = (|| {
            EmptyClipboard().map_err(|e| format!("cannot empty the clipboard: {}", e.message()))?;
            let memory: HGLOBAL = GlobalAlloc(GMEM_MOVEABLE, wide.len() * 2)
                .map_err(|e| format!("out of memory: {}", e.message()))?;
            let target = GlobalLock(memory) as *mut u16;
            if target.is_null() {
                let _ = GlobalFree(Some(memory));
                return Err("cannot lock memory for the clipboard".to_string());
            }
            std::ptr::copy_nonoverlapping(wide.as_ptr(), target, wide.len());
            let _ = GlobalUnlock(memory);
            // After success the clipboard owns the memory.
            if let Err(e) = SetClipboardData(CF_UNICODETEXT, Some(HANDLE(memory.0))) {
                let _ = GlobalFree(Some(memory));
                return Err(format!("cannot put the text on the clipboard: {}", e.message()));
            }
            Ok(())
        })();
        let _ = CloseClipboard();
        result
    }
}
