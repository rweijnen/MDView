//! Dropping a Markdown file on the standalone viewer's window (#54).
//!
//! An OLE drop target, so the cursor shows during the drag whether the drop
//! is possible: only one `.md` / `.markdown` file is accepted. The dropped
//! file is opened like File > Open (Mark of the Web check included), outside
//! the drop callback.

use std::cell::{Cell, RefCell};
use std::path::PathBuf;

use windows::Win32::Foundation::{HWND, LPARAM, POINTL, WPARAM};
use windows::Win32::System::Com::{DVASPECT_CONTENT, FORMATETC, IDataObject, TYMED_HGLOBAL};
use windows::Win32::System::Ole::{
    CF_HDROP, DROPEFFECT, DROPEFFECT_COPY, DROPEFFECT_NONE, IDropTarget, IDropTarget_Impl, RegisterDragDrop,
    ReleaseStgMedium,
};
use windows::Win32::System::SystemServices::MODIFIERKEYS_FLAGS;
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP};
use windows::Win32::UI::WindowsAndMessaging::PostMessageW;
use windows::core::{Ref, implement};

/// The one Markdown file in a drop, if that is what it is.
pub fn markdown_file(paths: &[PathBuf]) -> Option<&PathBuf> {
    match paths {
        [one] if crate::links::is_markdown(one) => Some(one),
        _ => None,
    }
}

/// The files in a drop (CF_HDROP), at their full length.
fn dropped_paths(data: &IDataObject) -> Vec<PathBuf> {
    let format = FORMATETC {
        cfFormat: CF_HDROP.0,
        ptd: std::ptr::null_mut(),
        dwAspect: DVASPECT_CONTENT.0,
        lindex: -1,
        tymed: TYMED_HGLOBAL.0 as u32,
    };
    let mut paths = Vec::new();
    unsafe {
        let Ok(mut medium) = data.GetData(&format) else { return paths };
        let hdrop = HDROP(medium.u.hGlobal.0);
        let count = DragQueryFileW(hdrop, u32::MAX, None);
        for i in 0..count {
            let len = DragQueryFileW(hdrop, i, None) as usize;
            let mut buffer = vec![0u16; len + 1];
            let copied = DragQueryFileW(hdrop, i, Some(&mut buffer)) as usize;
            paths.push(PathBuf::from(String::from_utf16_lossy(&buffer[..copied])));
        }
        ReleaseStgMedium(&mut medium);
    }
    paths
}

thread_local! {
    /// The file dropped, waiting for the window message that opens it.
    static DROPPED: RefCell<Option<PathBuf>> = const { RefCell::new(None) };
}

/// The dropped file, once (see `register`).
pub fn take_dropped() -> Option<PathBuf> {
    DROPPED.with(|d| d.borrow_mut().take())
}

#[implement(IDropTarget)]
struct DropTarget {
    window: HWND,
    message: u32,
    accept: Cell<bool>,
}

impl DropTarget {
    fn effect(&self, effect: *mut DROPEFFECT) {
        if !effect.is_null() {
            unsafe { *effect = if self.accept.get() { DROPEFFECT_COPY } else { DROPEFFECT_NONE } };
        }
    }
}

impl IDropTarget_Impl for DropTarget_Impl {
    fn DragEnter(
        &self,
        data: Ref<'_, IDataObject>,
        _keys: MODIFIERKEYS_FLAGS,
        _point: &POINTL,
        effect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let accept = data.as_ref().is_some_and(|d| markdown_file(&dropped_paths(d)).is_some());
        self.accept.set(accept);
        self.effect(effect);
        Ok(())
    }

    fn DragOver(&self, _keys: MODIFIERKEYS_FLAGS, _point: &POINTL, effect: *mut DROPEFFECT) -> windows::core::Result<()> {
        self.effect(effect);
        Ok(())
    }

    fn DragLeave(&self) -> windows::core::Result<()> {
        self.accept.set(false);
        Ok(())
    }

    fn Drop(
        &self,
        data: Ref<'_, IDataObject>,
        _keys: MODIFIERKEYS_FLAGS,
        _point: &POINTL,
        effect: *mut DROPEFFECT,
    ) -> windows::core::Result<()> {
        let file = data.as_ref().and_then(|d| markdown_file(&dropped_paths(d)).cloned());
        self.accept.set(file.is_some());
        self.effect(effect);
        if let Some(file) = file {
            DROPPED.with(|d| *d.borrow_mut() = Some(file));
            let _ = unsafe { PostMessageW(Some(self.window), self.message, WPARAM(0), LPARAM(0)) };
        }
        Ok(())
    }
}

/// Accept drops on `window`; a dropped Markdown file is announced with
/// `message` (then `take_dropped`). Needs OLE on this thread.
pub fn register(window: HWND, message: u32) -> Result<(), String> {
    let target: IDropTarget = DropTarget { window, message, accept: Cell::new(false) }.into();
    unsafe { RegisterDragDrop(window, &target) }.map_err(|e| e.message().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_only_one_markdown_file_is_accepted() {
        let p = |s: &str| PathBuf::from(s);
        assert_eq!(markdown_file(&[p(r"C:\a.md")]), Some(&p(r"C:\a.md")));
        assert_eq!(markdown_file(&[p(r"C:\a.MARKDOWN")]), Some(&p(r"C:\a.MARKDOWN")));
        assert_eq!(markdown_file(&[p(r"C:\a.txt")]), None);
        assert_eq!(markdown_file(&[p(r"C:\a.md"), p(r"C:\b.md")]), None, "several files");
        assert_eq!(markdown_file(&[]), None);
        let long = format!(r"C:\{}\a.md", "x".repeat(300));
        assert_eq!(markdown_file(&[p(&long)]), Some(&p(&long)), "longer than MAX_PATH");
    }
}
