//! Back and forward between documents (#46). Shared by the standalone viewer
//! and the TC plugin.
//!
//! WebView2 keeps the history of URLs, but MDView's document URLs don't name
//! the drive or share (they are relative to the current document's root).
//! So MDView keeps the documents shown, with their real paths, and on back /
//! forward makes the right one current before WebView2 asks for the page.

use std::path::{Path, PathBuf};

/// What kind of navigation WebView2 starts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    New,
    BackOrForward,
    Reload,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Entry {
    url: String,
    path: PathBuf,
}

#[derive(Debug, Default)]
pub struct History {
    entries: Vec<Entry>,
    pos: usize,
}

/// A URL without its `#heading`.
fn document_part(url: &str) -> &str {
    url.split('#').next().unwrap_or(url)
}

impl History {
    /// A document navigation starts. `resolve` maps a new URL to its file.
    /// Returns the document's path, which becomes the current document, or
    /// `None` when MDView doesn't know it (it is then refused elsewhere).
    pub fn navigate(&mut self, kind: Kind, url: &str, resolve: impl Fn(&str) -> Option<PathBuf>) -> Option<PathBuf> {
        let url = document_part(url);
        match kind {
            Kind::Reload => self.entries.get(self.pos).filter(|e| e.url == url).map(|e| e.path.clone()).or_else(|| resolve(url)),
            Kind::BackOrForward => {
                if self.pos > 0 && self.entries[self.pos - 1].url == url {
                    self.pos -= 1;
                } else if self.entries.get(self.pos + 1).is_some_and(|e| e.url == url) {
                    self.pos += 1;
                } else if let Some(i) = self.entries.iter().position(|e| e.url == url) {
                    // Several steps at once (a long press on Back).
                    self.pos = i;
                } else {
                    return self.push(url, resolve(url)?);
                }
                Some(self.entries[self.pos].path.clone())
            }
            Kind::New => {
                let path = resolve(url)?;
                self.push(url, path)
            }
        }
    }

    fn push(&mut self, url: &str, path: PathBuf) -> Option<PathBuf> {
        if self.entries.get(self.pos).is_some_and(|e| e.url == url && e.path == path) {
            return Some(path);
        }
        if !self.entries.is_empty() {
            self.entries.truncate(self.pos + 1);
        }
        self.entries.push(Entry { url: url.to_string(), path: path.clone() });
        self.pos = self.entries.len() - 1;
        Some(path)
    }

    /// The file name of the document before the current one, for "Back to ...".
    pub fn previous_name(&self) -> Option<String> {
        (self.pos > 0).then(|| file_name(&self.entries[self.pos - 1].path))
    }
}

fn file_name(path: &Path) -> String {
    path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_else(|| path.display().to_string())
}

/// JSON for `PostWebMessageAsJson`: whether the page can go back, and to what.
pub fn history_message(can_go_back: bool, previous: Option<&str>) -> String {
    format!(
        r#"{{"type":"history","back":{can_go_back},"previous":{}}}"#,
        crate::markdown::script_string(previous.unwrap_or(""))
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn on(drive: &'static str) -> impl Fn(&str) -> Option<PathBuf> {
        move |url: &str| url.strip_prefix("https://mdview.example/").map(|p| PathBuf::from(format!("{drive}\\{}", p.replace('/', "\\"))))
    }

    // Back to a document on another drive finds its real path again, although
    // its URL doesn't name the drive.
    #[test]
    fn test_back_and_forward_keep_the_real_paths() {
        let mut h = History::default();
        let a = "https://mdview.example/docs/a.md";
        let b = "https://mdview.example/notes/b.md";
        assert_eq!(h.navigate(Kind::New, a, on("C:")), Some(PathBuf::from(r"C:\docs\a.md")));
        assert_eq!(h.navigate(Kind::New, b, on("D:")), Some(PathBuf::from(r"D:\notes\b.md")));
        assert_eq!(h.previous_name().as_deref(), Some("a.md"));
        // Back: C:\docs\a.md, not D:\docs\a.md.
        assert_eq!(h.navigate(Kind::BackOrForward, a, on("D:")), Some(PathBuf::from(r"C:\docs\a.md")));
        assert_eq!(h.previous_name(), None);
        assert_eq!(h.navigate(Kind::BackOrForward, b, on("C:")), Some(PathBuf::from(r"D:\notes\b.md")));
        assert_eq!(h.navigate(Kind::Reload, b, on("C:")), Some(PathBuf::from(r"D:\notes\b.md")));
    }

    #[test]
    fn test_a_new_document_after_going_back_drops_the_forward_part() {
        let mut h = History::default();
        let u = |n: &str| format!("https://mdview.example/{n}.md");
        h.navigate(Kind::New, &u("a"), on("C:"));
        h.navigate(Kind::New, &u("b"), on("C:"));
        h.navigate(Kind::New, &u("c"), on("C:"));
        h.navigate(Kind::BackOrForward, &u("b"), on("C:"));
        h.navigate(Kind::New, &u("d"), on("C:"));
        assert_eq!(h.previous_name().as_deref(), Some("b.md"));
        assert_eq!(h.navigate(Kind::BackOrForward, &u("c"), on("C:")), Some(PathBuf::from(r"C:\c.md")), "unknown entry: resolved as new");
        // The same document again (a reload by navigation) is one entry.
        h.navigate(Kind::New, &u("c"), on("C:"));
        assert_eq!(h.previous_name().as_deref(), Some("d.md"));
        // A heading in the URL is not another document.
        h.navigate(Kind::New, &format!("{}#top", u("c")), on("C:"));
        assert_eq!(h.previous_name().as_deref(), Some("d.md"));
    }

    #[test]
    fn test_history_message_is_json() {
        assert_eq!(history_message(true, Some("a <b>.md")), r#"{"type":"history","back":true,"previous":"a \u003cb\u003e.md"}"#);
        assert_eq!(history_message(false, None), r#"{"type":"history","back":false,"previous":""}"#);
    }
}
