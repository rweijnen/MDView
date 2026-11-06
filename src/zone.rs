//! Mark of the Web: whether a document was downloaded from the internet (#24).
//!
//! Browsers, mail programs and unzip tools record where a file came from in
//! its `Zone.Identifier` stream. MDView does not show a marked document until
//! the user chooses to open it.

use std::ffi::OsString;
use std::path::Path;

/// A Mark of the Web that makes MDView ask before showing the document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Mark {
    /// Where the file was downloaded from, if the stream says so (http/https only).
    pub host_url: Option<String>,
}

/// The mark of `path`, if it is marked as coming from the internet (zone 3)
/// or from a restricted site (zone 4). A file without the stream, or on a
/// volume without named streams, is not marked. A stream that exists but
/// cannot be read or understood counts as marked: asking once too often is
/// harmless, showing a downloaded document without asking is not.
pub fn mark_of_the_web(path: &Path) -> Option<Mark> {
    if volume_has_named_streams(path) == Some(false) {
        return None;
    }
    match std::fs::read(zone_stream(path)) {
        Ok(bytes) => parse_zone_identifier(&bytes),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => None,
        Err(_) => Some(Mark { host_url: None }),
    }
}

/// Remove the mark, like Unblock in Explorer's file properties. A file
/// without a mark is fine.
pub fn remove_mark(path: &Path) -> Result<(), String> {
    match std::fs::remove_file(zone_stream(path)) {
        Ok(()) => Ok(()),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(e) => Err(format!("{}\n\n{e}", path.display())),
    }
}

fn zone_stream(path: &Path) -> OsString {
    let mut stream = path.as_os_str().to_owned();
    stream.push(":Zone.Identifier");
    stream
}

/// `None` if the volume could not be determined; the stream read decides then.
fn volume_has_named_streams(path: &Path) -> Option<bool> {
    use std::os::windows::ffi::OsStrExt;
    use windows::Win32::Storage::FileSystem::{GetVolumeInformationW, GetVolumePathNameW};
    use windows::Win32::System::SystemServices::FILE_NAMED_STREAMS;
    use windows::core::PCWSTR;

    let wide: Vec<u16> = path.as_os_str().encode_wide().chain(std::iter::once(0)).collect();
    let mut root = vec![0u16; wide.len().max(261)];
    unsafe { GetVolumePathNameW(PCWSTR(wide.as_ptr()), &mut root) }.ok()?;
    let mut flags = 0u32;
    unsafe { GetVolumeInformationW(PCWSTR(root.as_ptr()), None, None, None, Some(&mut flags), None) }
        .ok()?;
    Some(flags & FILE_NAMED_STREAMS != 0)
}

/// The mark described by the contents of a `Zone.Identifier` stream. Without
/// a readable `ZoneId` the file counts as marked.
fn parse_zone_identifier(bytes: &[u8]) -> Option<Mark> {
    let text = if let Some(utf16) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        let units: Vec<u16> = utf16
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        String::from_utf8_lossy(bytes).into_owned()
    };
    let value = |name: &str| {
        text.lines().find_map(|line| {
            let (key, value) = line.split_once('=')?;
            key.trim().eq_ignore_ascii_case(name).then(|| value.trim().to_string())
        })
    };
    let marked = match value("ZoneId").map(|v| v.parse::<u32>()) {
        Some(Ok(id)) => id == 3 || id == 4,
        Some(Err(_)) | None => true,
    };
    marked.then(|| Mark {
        host_url: value("HostUrl").filter(|url| {
            let lower = url.to_ascii_lowercase();
            lower.starts_with("https://") || lower.starts_with("http://")
        }),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn marked(bytes: &[u8]) -> bool {
        parse_zone_identifier(bytes).is_some()
    }

    #[test]
    fn test_zone_ids_that_mark_the_internet() {
        let marks = |id: &str| marked(format!("[ZoneTransfer]\r\nZoneId={id}\r\n").as_bytes());
        assert!(marks("3"), "internet");
        assert!(marks("4"), "restricted sites");
        assert!(!marks("0"), "local computer");
        assert!(!marks("1"), "local intranet");
        assert!(!marks("2"), "trusted sites");
    }

    #[test]
    fn test_zone_identifier_variants_are_understood() {
        assert!(marked(b"[ZoneTransfer]\nzoneid = 3\nReferrerUrl=https://example.com/\n"));
        assert!(!marked(b"[ZoneTransfer]\r\n  ZoneId=2  \r\n"));
        let utf16: Vec<u8> = [0xFF, 0xFE]
            .into_iter()
            .chain("[ZoneTransfer]\r\nZoneId=3\r\n".encode_utf16().flat_map(u16::to_le_bytes))
            .collect();
        assert!(marked(&utf16), "UTF-16 with BOM");
    }

    #[test]
    fn test_unreadable_zone_identifier_counts_as_marked() {
        assert!(marked(b""), "empty stream");
        assert!(marked(b"\x00\x01garbage"), "garbage");
        assert!(marked(b"[ZoneTransfer]\r\nHostUrl=https://example.com/\r\n"), "no ZoneId");
        assert!(marked(b"[ZoneTransfer]\r\nZoneId=three\r\n"), "ZoneId not a number");
    }

    #[test]
    fn test_host_url_is_kept_only_for_web_addresses() {
        let host = |s: &str| parse_zone_identifier(s.as_bytes()).and_then(|m| m.host_url);
        assert_eq!(
            host("[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=https://example.com/a.md\r\n").as_deref(),
            Some("https://example.com/a.md")
        );
        assert_eq!(host("[ZoneTransfer]\r\nZoneId=3\r\nhosturl = HTTP://x/\r\n").as_deref(), Some("HTTP://x/"));
        assert_eq!(host("[ZoneTransfer]\r\nZoneId=3\r\n"), None, "no HostUrl");
        assert_eq!(host("[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=about:internet\r\n"), None, "not a web address");
        assert_eq!(host("[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=javascript:x()\r\n"), None);
    }

    fn temp_doc(name: &str) -> (std::path::PathBuf, std::path::PathBuf) {
        let dir = std::env::temp_dir().join(format!("mdview-zone-{name}-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("doc.md");
        std::fs::write(&file, "# doc").unwrap();
        (dir, file)
    }

    #[test]
    fn test_mark_of_the_web_is_read_from_the_file() {
        let (dir, file) = temp_doc("read");
        let stream = dir.join("doc.md:Zone.Identifier");
        assert_eq!(mark_of_the_web(&file), None, "no stream");
        std::fs::write(&stream, "[ZoneTransfer]\r\nZoneId=3\r\nHostUrl=https://example.com/doc.md\r\n").unwrap();
        assert_eq!(
            mark_of_the_web(&file),
            Some(Mark { host_url: Some("https://example.com/doc.md".to_string()) })
        );
        std::fs::write(&stream, "[ZoneTransfer]\r\nZoneId=2\r\n").unwrap();
        assert_eq!(mark_of_the_web(&file), None, "zone 2");
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn test_remove_mark_unblocks_the_file() {
        let (dir, file) = temp_doc("remove");
        std::fs::write(dir.join("doc.md:Zone.Identifier"), "[ZoneTransfer]\r\nZoneId=3\r\n").unwrap();
        assert!(mark_of_the_web(&file).is_some());
        assert_eq!(remove_mark(&file), Ok(()));
        assert_eq!(mark_of_the_web(&file), None, "mark removed");
        assert_eq!(std::fs::read_to_string(&file).unwrap(), "# doc", "the document itself is untouched");
        assert_eq!(remove_mark(&file), Ok(()), "no mark: nothing to do");
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn test_remove_mark_says_why_it_failed() {
        use std::os::windows::fs::OpenOptionsExt;
        let (dir, file) = temp_doc("fail");
        let stream = dir.join("doc.md:Zone.Identifier");
        std::fs::write(&stream, "[ZoneTransfer]\r\nZoneId=3\r\n").unwrap();
        // Another program has the stream open without sharing it.
        let held = std::fs::OpenOptions::new().read(true).share_mode(0).open(&stream).unwrap();
        let result = remove_mark(&file);
        drop(held);
        let err = result.expect_err("a mark that is in use cannot be removed");
        assert!(err.contains(&file.display().to_string()), "names the file: {err}");
        assert!(mark_of_the_web(&file).is_some(), "still marked");
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
