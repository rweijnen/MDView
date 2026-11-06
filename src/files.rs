//! Reading the user's files: always read-only, and never in the way of other
//! programs. MDView opens a file only to read it, shares it for reading,
//! writing and deleting, and closes it again straight away.

use std::fs::File;
use std::io::Read;
use std::path::Path;

/// Open `path` for reading only. Other programs can keep reading, writing,
/// renaming and deleting the file while MDView has it open.
pub fn open_read_only(path: &Path) -> std::io::Result<File> {
    use std::os::windows::fs::OpenOptionsExt;
    use windows::Win32::Storage::FileSystem::{FILE_SHARE_DELETE, FILE_SHARE_READ, FILE_SHARE_WRITE};
    std::fs::OpenOptions::new()
        .read(true)
        .share_mode((FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE).0)
        .open(path)
}

/// The whole file, read-only. The error says which file and why.
pub fn read_bytes(path: &Path) -> Result<Vec<u8>, String> {
    let mut bytes = Vec::new();
    open_read_only(path)
        .and_then(|mut file| file.read_to_end(&mut bytes))
        .map_err(|e| e.to_string())?;
    Ok(bytes)
}

/// The whole file as text, read-only.
pub fn read_text(path: &Path) -> Result<String, String> {
    String::from_utf8(read_bytes(path)?).map_err(|_| "The file is not valid UTF-8 text.".to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_files_are_opened_read_only_and_never_in_the_way() {
        let dir = std::env::temp_dir().join(format!("mdview-files-test-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("doc.md");
        std::fs::write(&file, "# doc").unwrap();

        let mut open = open_read_only(&file).unwrap();
        // MDView's handle can only read.
        use std::io::Write;
        assert!(open.write_all(b"x").is_err(), "MDView's handle cannot write");
        // Another program can still read, write, rename and delete the file.
        assert_eq!(std::fs::read_to_string(&file).unwrap(), "# doc");
        std::fs::write(&file, "# changed").unwrap();
        let renamed = dir.join("renamed.md");
        std::fs::rename(&file, &renamed).unwrap();
        std::fs::remove_file(&renamed).unwrap();
        let mut text = String::new();
        open.read_to_string(&mut text).unwrap();
        assert_eq!(text, "# changed", "MDView reads what is there");
        drop(open);
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn test_read_errors_say_why() {
        let missing = std::env::temp_dir().join("mdview-files-test-missing.md");
        assert!(read_text(&missing).is_err());
        let dir = std::env::temp_dir().join(format!("mdview-files-utf8-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("latin1.md");
        std::fs::write(&file, b"caf\xe9").unwrap();
        assert_eq!(read_text(&file), Err("The file is not valid UTF-8 text.".to_string()));
        std::fs::remove_dir_all(&dir).unwrap();
    }
}
