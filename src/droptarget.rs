//! Dropping a Markdown file on the standalone viewer's window (#54).
//!
//! Two ways in, both opened like File > Open (Mark of the Web check
//! included), outside the callback that reported them:
//! - On the window's edge: an OLE drop target, so the cursor shows whether the
//!   drop is possible (only one `.md` / `.markdown` file).
//! - On the page: WebView2's own drop handling takes the drag there (the OLE
//!   target is not asked), and reports the file as a new-window request for
//!   its `file:` URL, which the host refuses and hands to `page_drop`. The cursor
//!   shows "copy" for any file; anything but Markdown gets a message.

use std::cell::{Cell, RefCell};
use std::path::{Path, PathBuf};

use windows::Win32::Foundation::{HWND, POINTL};
use windows::Win32::System::Com::{DVASPECT_CONTENT, FORMATETC, IDataObject, TYMED_HGLOBAL};
use windows::Win32::System::Ole::{
    CF_HDROP, DROPEFFECT, DROPEFFECT_COPY, DROPEFFECT_NONE, IDropTarget, IDropTarget_Impl, RegisterDragDrop,
    ReleaseStgMedium,
};
use windows::Win32::System::SystemServices::MODIFIERKEYS_FLAGS;
use windows::Win32::UI::Shell::{DragQueryFileW, HDROP};
use windows::Win32::UI::WindowsAndMessaging::SetTimer;
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

/// What was dropped.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Dropped {
    /// One Markdown file: open it.
    Markdown(PathBuf),
    /// Something else, as the user would recognise it: say MDView can't open it.
    NotMarkdown(String),
}

/// The new-window request WebView2 makes for a file dropped on the page:
/// its `file:` URL. `None` when `url` is not a `file:` URL. A document can ask
/// for such a window too (`window.open`; WebView2 can call it user-initiated,
/// see S5); the worst it gets is that MDView opens another local Markdown
/// file, as a clicked link can.
pub fn page_drop(url: &str) -> Option<Dropped> {
    if !url.get(..5).is_some_and(|s| s.eq_ignore_ascii_case("file:")) {
        return None;
    }
    Some(match crate::vpath::file_url_to_path(url) {
        Some(path) if crate::links::is_markdown(&path) => Dropped::Markdown(path),
        Some(path) => Dropped::NotMarkdown(path.display().to_string()),
        None => Dropped::NotMarkdown(url.to_string()),
    })
}

/// What MDView does with a drop.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Outcome {
    Open(PathBuf),
    /// On another share or a network drive: only after the user agrees, as a
    /// link there (opening it connects to `location`; a document can ask for
    /// it without a drop, see `page_drop`).
    AskNetwork { path: PathBuf, location: String },
    /// Nothing is opened; the user is told why.
    Tell(String),
}

/// One drop can arrive as several reports (WebView2 asks for a window per
/// dropped file, ~2 ms apart): reports within this time are one drop.
pub const COLLECT_MS: u32 = 150;

/// What one drop means: one Markdown file is opened (after a question when
/// `network_location` names a network location for it); anything else is
/// explained, never ignored.
pub fn outcome(drops: &[Dropped], network_location: impl Fn(&Path) -> Option<String>) -> Option<Outcome> {
    match drops {
        [] => None,
        [Dropped::Markdown(path)] => Some(match network_location(path) {
            Some(location) => Outcome::AskNetwork { path: path.clone(), location },
            None => Outcome::Open(path.clone()),
        }),
        [Dropped::NotMarkdown(what)] => {
            Some(Outcome::Tell(format!("MDView only opens Markdown files (.md, .markdown):\n\n{what}")))
        }
        _ => Some(Outcome::Tell("Drop one Markdown file at a time.".to_string())),
    }
}

thread_local! {
    /// The reports of the drop being collected (see `announce`).
    static DROPPED: RefCell<Vec<Dropped>> = const { RefCell::new(Vec::new()) };
}

/// Add a report to the drop being collected; the first one starts `timer` on
/// `window`, after which the window procedure calls `take_drop`.
pub fn announce(window: HWND, timer: usize, dropped: Dropped) {
    let first = DROPPED.with(|d| {
        let mut all = d.borrow_mut();
        all.push(dropped);
        all.len() == 1
    });
    if first {
        unsafe { SetTimer(Some(window), timer, COLLECT_MS, None) };
    }
}

/// The collected drop, once (see `announce`).
pub fn take_drop(network_location: impl Fn(&Path) -> Option<String>) -> Option<Outcome> {
    let drops = DROPPED.with(|d| std::mem::take(&mut *d.borrow_mut()));
    outcome(&drops, network_location)
}

#[implement(IDropTarget)]
struct DropTarget {
    window: HWND,
    timer: usize,
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
            announce(self.window, self.timer, Dropped::Markdown(file));
        }
        Ok(())
    }
}

/// Accept drops on `window`; a dropped Markdown file is announced with
/// `timer` (then `take_drop`). Needs OLE on this thread.
pub fn register(window: HWND, timer: usize) -> Result<(), String> {
    let target: IDropTarget = DropTarget { window, timer, accept: Cell::new(false) }.into();
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

    // #54: one drop is one Markdown file, or MDView says why not.
    #[test]
    fn test_what_a_drop_does() {
        let md = |s: &str| Dropped::Markdown(PathBuf::from(s));
        let local = |_: &Path| None;
        assert_eq!(outcome(&[], local), None);
        assert_eq!(outcome(&[md(r"C:\a.md")], local), Some(Outcome::Open(PathBuf::from(r"C:\a.md"))));
        assert_eq!(
            outcome(&[Dropped::NotMarkdown(r"C:\a.json".into())], local),
            Some(Outcome::Tell("MDView only opens Markdown files (.md, .markdown):\n\nC:\\a.json".into()))
        );
        // Two files dropped together (WebView2 reports each): not the first
        // or last one silently, but a message.
        let several = Some(Outcome::Tell("Drop one Markdown file at a time.".into()));
        assert_eq!(outcome(&[md(r"C:\a.md"), md(r"C:\b.md")], local), several);
        assert_eq!(outcome(&[md(r"C:\a.md"), Dropped::NotMarkdown(r"C:\b.txt".into())], local), several);
        // On another share: asked first, never opened by itself (a document
        // can ask for such a "drop" too).
        let share = |p: &Path| p.to_string_lossy().starts_with(r"\\evil").then(|| r"\\evil\share".to_string());
        assert_eq!(
            outcome(&[md(r"\\evil\share\x.md")], share),
            Some(Outcome::AskNetwork { path: PathBuf::from(r"\\evil\share\x.md"), location: r"\\evil\share".into() })
        );
        assert_eq!(outcome(&[md(r"C:\a.md")], share), Some(Outcome::Open(PathBuf::from(r"C:\a.md"))));
    }

    // #54: a file dropped on the page arrives as its file: URL.
    #[test]
    fn test_a_file_dropped_on_the_page() {
        let p = |s: &str| PathBuf::from(s);
        assert_eq!(page_drop("file:///C:/docs/a%20b.md"), Some(Dropped::Markdown(p(r"C:\docs\a b.md"))));
        assert_eq!(page_drop("FILE:///C:/docs/x.MARKDOWN"), Some(Dropped::Markdown(p(r"C:\docs\x.MARKDOWN"))));
        assert_eq!(page_drop("file://server/share/n.md"), Some(Dropped::Markdown(p(r"\\server\share\n.md"))));
        assert_eq!(page_drop("file:///C:/docs/notes.txt"), Some(Dropped::NotMarkdown(r"C:\docs\notes.txt".into())));
        assert_eq!(page_drop("file:///C:/tools/setup.exe"), Some(Dropped::NotMarkdown(r"C:\tools\setup.exe".into())));
        // Device paths are never opened, not even as Markdown.
        assert!(matches!(page_drop("file:////./pipe/p.md"), Some(Dropped::NotMarkdown(_))));
        // Not a file: URL, so not a drop: the normal navigation rules apply.
        for url in ["https://mdview.example/C/a.md", "about:blank", "data:text/html,x", "https://example.com/a.md"] {
            assert_eq!(page_drop(url), None, "{url}");
        }
    }
}
