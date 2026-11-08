//! Reload the shown document when its file changes on disk (#17). Shared by
//! the standalone viewer and the TC plugin.
//!
//! A timer asks once a second; only the file's write time and size are read
//! (no content, no change notifications, which are unreliable on network
//! shares). A change is acted on once the file has stopped changing, so a
//! document is not reloaded half written.

use std::path::{Path, PathBuf};

/// How often the file is checked.
pub const INTERVAL_MS: u32 = 1000;

/// What identifies a version of the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Stamp {
    pub written: u64,
    pub size: u64,
}

/// The file's stamp now; `None` while it doesn't exist or can't be read.
pub fn stamp(path: &Path) -> Option<Stamp> {
    let meta = std::fs::metadata(path).ok()?;
    let written = meta
        .modified()
        .ok()?
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64)
        .unwrap_or(0);
    Some(Stamp { written, size: meta.len() })
}

/// The file being watched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Watch {
    pub path: PathBuf,
    shown: Option<Stamp>,
    /// Seen at the last tick, different from what is shown.
    pending: Option<Option<Stamp>>,
}

impl Watch {
    /// Start watching `path`, as it is now (the version being shown).
    pub fn new(path: &Path, now: Option<Stamp>) -> Self {
        Watch { path: path.to_path_buf(), shown: now, pending: None }
    }

    /// One timer tick with the file's stamp now. Returns true when the page
    /// should be reloaded: the file changed and has been the same since the
    /// last tick. A file that disappeared is reloaded once it is back.
    pub fn tick(&mut self, now: Option<Stamp>) -> bool {
        if now == self.shown {
            self.pending = None;
            return false;
        }
        if self.pending == Some(now) {
            self.shown = now;
            self.pending = None;
            return now.is_some();
        }
        self.pending = Some(now);
        false
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(written: u64, size: u64) -> Option<Stamp> {
        Some(Stamp { written, size })
    }

    #[test]
    fn test_a_change_reloads_once_the_file_is_stable() {
        let mut w = Watch::new(Path::new("a.md"), s(1, 10));
        assert!(!w.tick(s(1, 10)), "unchanged");
        assert!(!w.tick(s(2, 12)), "just changed: wait for it to settle");
        assert!(!w.tick(s(3, 14)), "still being written");
        assert!(w.tick(s(3, 14)), "stable: reload");
        assert!(!w.tick(s(3, 14)), "reloaded already");
    }

    #[test]
    fn test_a_file_that_disappears_is_reloaded_when_it_is_back() {
        let mut w = Watch::new(Path::new("a.md"), s(1, 10));
        assert!(!w.tick(None));
        assert!(!w.tick(None), "gone: nothing to show, keep the page");
        assert!(!w.tick(s(5, 10)));
        assert!(w.tick(s(5, 10)), "back and stable: reload");
    }

    #[test]
    fn test_stamp_reads_size_and_time_only() {
        let dir = std::env::temp_dir().join(format!("mdview-watch-test-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("a.md");
        std::fs::write(&file, "one").unwrap();
        let first = stamp(&file).unwrap();
        assert_eq!(first.size, 3);
        std::fs::write(&file, "three").unwrap();
        assert_eq!(stamp(&file).unwrap().size, 5);
        assert!(stamp(&dir.join("missing.md")).is_none());
        let _ = std::fs::remove_dir_all(&dir);
    }
}
