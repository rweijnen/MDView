//! MDView's virtual URLs: how files are named on the page, and how a URL the
//! page requests is mapped back to a file. Shared by the standalone viewer and
//! the Total Commander plugin, so both apply exactly the same rules.
//!
//! - `https://mdview.example/<path>`: a file on the document's own drive or
//!   share (the path from its root). Relative links resolve through ordinary
//!   URL resolution.
//! - `https://drive-<x>.mdview.example/<path>`: a file on another drive.
//! - `https://unc.mdview.example/<server>/<share>/<path>`: a file on another
//!   network share.
//!
//! A URL comes from an untrusted document. It maps to a file only if every
//! part of its path is a plain file or folder name: encoded separators, drive
//! letters, `..`, devices (`CON`, `NUL`, ...), pipes, `\\?\` and alternate
//! data streams never become a path. Paths are built with `\` only (#49).

use std::path::{Component, Path, PathBuf, Prefix};

const DOCUMENT_HOST: &str = "mdview.example";
const UNC_HOST: &str = "unc.mdview.example";

/// The root a file lives under: a drive letter or a network share.
#[derive(Debug, Clone)]
pub enum Root {
    Disk(char),
    Unc { server: String, share: String },
}

impl Root {
    /// The root of an absolute (or, for the document itself, relative) path.
    pub fn of(path: &Path) -> Option<Root> {
        split(path).map(|(root, _)| root)
    }

    /// As Windows writes it: `C:\` or `\\server\share\`.
    pub fn display(&self) -> String {
        match self {
            Root::Disk(letter) => format!("{letter}:\\"),
            Root::Unc { server, share } => format!("\\\\{server}\\{share}\\"),
        }
    }

    /// The location as shown to the user: `Z:` or `\\server\share`.
    pub fn name(&self) -> String {
        match self {
            Root::Disk(letter) => format!("{letter}:"),
            Root::Unc { server, share } => format!("\\\\{server}\\{share}"),
        }
    }

    /// Windows compares drive letters, server and share names without case.
    pub fn same(&self, other: &Root) -> bool {
        match (self, other) {
            (Root::Disk(a), Root::Disk(b)) => a.eq_ignore_ascii_case(b),
            (Root::Unc { server: s1, share: h1 }, Root::Unc { server: s2, share: h2 }) => {
                s1.to_lowercase() == s2.to_lowercase() && h1.to_lowercase() == h2.to_lowercase()
            }
            _ => false,
        }
    }
}

/// Whether `name` is an ordinary file or folder name that Windows opens as
/// itself. Device names and names Windows would change are refused.
fn plain_name(name: &str) -> bool {
    !name.is_empty()
        && name != "."
        && name != ".."
        && !name.ends_with(['.', ' '])
        && !name.chars().any(|c| {
            (c as u32) < 0x20 || matches!(c, '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|')
        })
        && !is_device_name(name)
}

/// `CON`, `nul.txt`, `COM1.md`, `LPT²`, ...: Windows opens the device, not a file.
fn is_device_name(name: &str) -> bool {
    let base = name.split('.').next().unwrap_or_default().trim_end_matches(' ').to_ascii_uppercase();
    if matches!(base.as_str(), "CON" | "PRN" | "AUX" | "NUL" | "CONIN$" | "CONOUT$") {
        return true;
    }
    let mut chars = base.chars();
    let prefix: String = chars.by_ref().take(3).collect();
    let rest: Vec<char> = chars.collect();
    (prefix == "COM" || prefix == "LPT")
        && rest.len() == 1
        && matches!(rest[0], '0'..='9' | '\u{00B9}' | '\u{00B2}' | '\u{00B3}')
}

/// A path split into its root and plain names. `..` goes up, but never past
/// the root. Anything that is not a plain name refuses the whole path.
fn split(path: &Path) -> Option<(Root, Vec<String>)> {
    let absolute;
    let path = if path.is_absolute() {
        path
    } else {
        absolute = std::path::absolute(path).ok()?;
        &absolute
    };
    let mut components = path.components();
    let root = match components.next()? {
        Component::Prefix(prefix) => match prefix.kind() {
            Prefix::Disk(letter) | Prefix::VerbatimDisk(letter) => {
                Root::Disk((letter as char).to_ascii_uppercase())
            }
            Prefix::UNC(server, share) | Prefix::VerbatimUNC(server, share) => Root::Unc {
                server: server.to_str()?.to_string(),
                share: share.to_str()?.to_string(),
            },
            _ => return None,
        },
        _ => return None,
    };
    if let Root::Unc { server, share } = &root {
        if !plain_name(server) || !plain_name(share) {
            return None;
        }
    }
    if components.next()? != Component::RootDir {
        return None;
    }
    let mut parts = Vec::new();
    for component in components {
        match component {
            Component::Normal(name) => {
                let name = name.to_str()?;
                if !plain_name(name) {
                    return None;
                }
                parts.push(name.to_string());
            }
            Component::CurDir => {}
            Component::ParentDir => {
                parts.pop();
            }
            _ => return None,
        }
    }
    Some((root, parts))
}

fn percent_encode(name: &str) -> String {
    let mut encoded = String::with_capacity(name.len());
    for byte in name.bytes() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => encoded.push(byte as char),
            _ => encoded.push_str(&format!("%{byte:02X}")),
        }
    }
    encoded
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

/// Decoded once: `%255C` is the name `%5C`, never a `\`. Not valid UTF-8: None.
pub fn percent_decode(input: &str) -> Option<String> {
    let bytes = input.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(hi), Some(lo)) = (hex(bytes[i + 1]), hex(bytes[i + 2])) {
                decoded.push(hi << 4 | lo);
                i += 3;
                continue;
            }
        }
        decoded.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(decoded).ok()
}

/// The URL that names `path` on a page showing a document under `doc_root`.
pub fn url_for(path: &Path, doc_root: &Root) -> Option<String> {
    let (root, parts) = split(path)?;
    if parts.is_empty() {
        return None;
    }
    let base = if root.same(doc_root) {
        format!("https://{DOCUMENT_HOST}/")
    } else {
        match &root {
            Root::Disk(letter) => format!("https://drive-{}.{DOCUMENT_HOST}/", letter.to_ascii_lowercase()),
            Root::Unc { server, share } => {
                format!("https://{UNC_HOST}/{}/{}/", percent_encode(server), percent_encode(share))
            }
        }
    };
    let parts: Vec<String> = parts.iter().map(|p| percent_encode(p)).collect();
    Some(base + &parts.join("/"))
}

/// The URL a document itself is shown at.
pub fn document_url(document: &Path) -> Option<String> {
    url_for(document, &Root::of(document)?)
}

/// `(host, path)` of one of MDView's own URLs, without query and fragment.
fn virtual_parts(url: &str) -> Option<(String, &str)> {
    let rest = url.strip_prefix("https://").or_else(|| url.strip_prefix("http://"))?;
    let (host, path) = rest.split_once('/')?;
    let host = host.to_ascii_lowercase();
    if host != DOCUMENT_HOST && !host.ends_with(".mdview.example") {
        return None;
    }
    Some((host, path.split(['?', '#']).next().unwrap_or_default()))
}

/// Any of MDView's own URLs, whether or not it maps to a file.
pub fn is_virtual_url(url: &str) -> bool {
    virtual_parts(url).is_some()
}

/// A URL on the document's own root (`https://mdview.example/...`). Only
/// these can be a page; the other hosts only ever serve files to a page.
pub fn is_document_origin(url: &str) -> bool {
    virtual_parts(url).is_some_and(|(host, _)| host == DOCUMENT_HOST)
}

/// The file a URL from the page stands for, while `document` is shown.
/// None if it is not one of MDView's URLs or does not name a plain file.
pub fn resolve(url: &str, document: &Path) -> Option<PathBuf> {
    let (host, path) = virtual_parts(url)?;
    let mut names = path
        .split('/')
        .map(percent_decode)
        .collect::<Option<Vec<String>>>()?;
    if names.iter().any(|name| !plain_name(name)) {
        return None;
    }
    let root = if host == DOCUMENT_HOST {
        Root::of(document)?
    } else if host == UNC_HOST {
        if names.len() < 3 {
            return None;
        }
        let server = names.remove(0);
        let share = names.remove(0);
        Root::Unc { server, share }
    } else {
        let label = host.strip_suffix(".mdview.example")?.strip_prefix("drive-")?;
        let mut letters = label.chars();
        match (letters.next(), letters.next()) {
            (Some(letter), None) if letter.is_ascii_lowercase() => Root::Disk(letter.to_ascii_uppercase()),
            _ => return None,
        }
    };
    if names.is_empty() {
        return None;
    }
    Some(PathBuf::from(root.display() + &names.join("\\")))
}

/// What a drive letter stands for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DriveKind {
    Local,
    Network,
    /// No such drive: nothing to connect to.
    Missing,
}

/// From `GetDriveTypeW` and `QueryDosDeviceW`: a mapped network drive, or a
/// drive letter that points at a share (`subst`/`DefineDosDevice`), is a
/// network location. A drive type Windows does not know counts as network,
/// so nothing unknown is opened without asking.
pub fn drive_kind_from(drive_type: u32, dos_target: Option<&str>) -> DriveKind {
    const DRIVE_UNKNOWN: u32 = 0;
    const DRIVE_NO_ROOT_DIR: u32 = 1;
    const DRIVE_REMOTE: u32 = 4;
    let target = dos_target.unwrap_or_default().to_ascii_lowercase();
    if drive_type == DRIVE_REMOTE
        || drive_type == DRIVE_UNKNOWN
        || target.starts_with(r"\??\unc\")
        || target.starts_with(r"\device\lanmanredirector")
        || target.starts_with(r"\device\mup")
        || target.starts_with(r"\device\webdavredirector")
    {
        DriveKind::Network
    } else if drive_type == DRIVE_NO_ROOT_DIR {
        DriveKind::Missing
    } else {
        DriveKind::Local
    }
}

/// The real probe. Asks Windows what the drive letter is; this does not
/// connect to anything.
pub fn drive_kind(letter: char) -> DriveKind {
    use windows::Win32::Storage::FileSystem::{GetDriveTypeW, QueryDosDeviceW};
    use windows::core::PCWSTR;
    let root: Vec<u16> = format!("{letter}:\\\0").encode_utf16().collect();
    let device: Vec<u16> = format!("{letter}:\0").encode_utf16().collect();
    let drive_type = unsafe { GetDriveTypeW(PCWSTR(root.as_ptr())) };
    let mut buffer = vec![0u16; 1024];
    let written = unsafe { QueryDosDeviceW(PCWSTR(device.as_ptr()), Some(&mut buffer)) } as usize;
    let target = (written > 0).then(|| {
        let end = buffer[..written.min(buffer.len())].iter().position(|&c| c == 0).unwrap_or(written);
        String::from_utf16_lossy(&buffer[..end])
    });
    drive_kind_from(drive_type, target.as_deref())
}

/// Where a file is, seen from the document being shown.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Location {
    /// The document's own drive or share.
    SameRoot,
    /// Another local drive.
    Local,
    /// Another share or a network drive, named as shown to the user.
    Network(String),
}

/// Decided from the path text and `probe` alone: never touches the network.
pub fn location(path: &Path, document: &Path, probe: impl Fn(char) -> DriveKind) -> Option<Location> {
    let root = Root::of(path)?;
    if root.same(&Root::of(document)?) {
        return Some(Location::SameRoot);
    }
    Some(match &root {
        Root::Unc { .. } => Location::Network(root.name()),
        Root::Disk(letter) => match probe(*letter) {
            DriveKind::Network => Location::Network(root.name()),
            DriveKind::Local | DriveKind::Missing => Location::Local,
        },
    })
}

/// `\\?\...` and `\\.\...` name devices and raw paths, never a document's file.
fn special_path(path: &str) -> bool {
    path.starts_with(r"\\?\") || path.starts_with(r"\\.\") || path.starts_with(r"\??\")
}

/// `file:///C:/dir/a%20b.txt` or `file://server/share/x` to a Windows path.
pub fn file_url_to_path(url: &str) -> Option<PathBuf> {
    let rest = url.get(..5).filter(|s| s.eq_ignore_ascii_case("file:")).and(url.get(5..))?;
    let rest = rest.split(['?', '#']).next().unwrap_or_default();
    let decoded = percent_decode(rest)?;
    let path = if let Some(local) = decoded.strip_prefix("///") {
        local.replace('/', "\\")
    } else if let Some(unc) = decoded.strip_prefix("//") {
        let unc = unc.replace('/', "\\");
        match unc.split_once('\\') {
            // file://localhost/C:/x is C:\x
            Some((host, local)) if host.eq_ignore_ascii_case("localhost") && local.get(1..2) == Some(":") => {
                local.to_string()
            }
            _ => format!("\\\\{unc}"),
        }
    } else {
        return None;
    };
    (drive_or_unc(&path) && !special_path(&path)).then(|| PathBuf::from(path))
}

/// `X:\...` or `\\...`: an absolute path with a drive letter or a share.
fn drive_or_unc(path: &str) -> bool {
    let bytes = path.as_bytes();
    (bytes.len() >= 3 && bytes[0].is_ascii_alphabetic() && bytes[1] == b':' && matches!(bytes[2], b'\\' | b'/'))
        || path.starts_with(r"\\")
}

/// An absolute path written in a document's link or image: `D:\x.png`,
/// `D:/x.png`, `file:///D:/x.png`, `\\server\share\x.png` or
/// `file://server/share/x.png`. Relative paths, web links, `#anchors` and
/// `//host` (which browsers read as a web address) are None.
pub fn absolute_link_target(dest: &str) -> Option<PathBuf> {
    let dest = dest.trim();
    if dest.get(..5).is_some_and(|s| s.eq_ignore_ascii_case("file:")) {
        return file_url_to_path(dest);
    }
    if !drive_or_unc(dest) {
        return None;
    }
    let path = dest.replace('/', "\\");
    (!special_path(&path)).then(|| PathBuf::from(path))
}

#[cfg(test)]
mod tests {
    use super::*;

    const DOC: &str = r"C:\docs\note.md";

    fn resolves(url: &str, document: &str) -> Option<String> {
        resolve(url, Path::new(document)).map(|p| p.to_string_lossy().into_owned())
    }

    #[test]
    fn test_document_and_its_files_use_the_plain_host_and_backslashes() {
        assert_eq!(document_url(Path::new(DOC)).as_deref(), Some("https://mdview.example/docs/note.md"));
        assert_eq!(resolves("https://mdview.example/docs/img/a.png?v=2#x", DOC).as_deref(), Some(r"C:\docs\img\a.png"));
        let share = r"\\server\share\docs\note.md";
        assert_eq!(document_url(Path::new(share)).as_deref(), Some("https://mdview.example/docs/note.md"));
        assert_eq!(resolves("https://mdview.example/docs/a.png", share).as_deref(), Some(r"\\server\share\docs\a.png"));
    }

    #[test]
    fn test_names_with_spaces_and_non_ascii_round_trip() {
        let image = r"C:\Dokumenty\poznámky\snímek obrazovky.png";
        let url = url_for(Path::new(image), &Root::Disk('C')).unwrap();
        assert_eq!(url, "https://mdview.example/Dokumenty/pozn%C3%A1mky/sn%C3%ADmek%20obrazovky.png");
        assert_eq!(resolves(&url, DOC).as_deref(), Some(image));
        // A '%' in a real name survives.
        let percent = r"C:\docs\100%.png";
        assert_eq!(resolves(&url_for(Path::new(percent), &Root::Disk('C')).unwrap(), DOC).as_deref(), Some(percent));
    }

    #[test]
    fn test_other_drives_and_shares_get_their_own_host() {
        let doc_root = Root::Disk('C');
        for (path, url) in [
            (r"D:\pictures\x.png", "https://drive-d.mdview.example/pictures/x.png"),
            (r"d:\pictures\x.png", "https://drive-d.mdview.example/pictures/x.png"),
            (r"\\server\share\img\x.png", "https://unc.mdview.example/server/share/img/x.png"),
            (r"C:\other\x.png", "https://mdview.example/other/x.png"),
        ] {
            assert_eq!(url_for(Path::new(path), &doc_root).as_deref(), Some(url), "{path}");
            assert_eq!(resolves(url, DOC).as_deref(), Some(path.replace("d:", "D:").as_str()), "{url}");
        }
        // A document on a share names files on that share with the plain host.
        let share_root = Root::of(Path::new(r"\\SERVER\Share\docs\note.md")).unwrap();
        assert_eq!(
            url_for(Path::new(r"\\server\share\img\x.png"), &share_root).as_deref(),
            Some("https://mdview.example/img/x.png")
        );
    }

    #[test]
    fn test_urls_cannot_escape_to_other_paths() {
        for url in [
            "https://mdview.example/%5C%5Cserver%5Cshare%5Cx.png",
            "https://mdview.example/%2F%2Fserver%2Fshare%2Fx.png",
            "https://mdview.example/Q%3A%5Cx.png",
            "https://mdview.example/Q%3A/x.png",
            "https://mdview.example/docs/..%5C..%5Cx.png",
            "https://mdview.example/docs/%2E%2E/x.png",
            "https://mdview.example/docs/%2E/x.png",
            "https://mdview.example/%5C%5C.%5Cpipe%5Cp",
            "https://mdview.example/%5C%5C%3F%5CC%3A%5Cx.png",
            "https://mdview.example/docs/a.png:secret",
            "https://mdview.example/docs/a.png%3Asecret",
            "https://mdview.example/docs/CON",
            "https://mdview.example/docs/nul.txt",
            "https://mdview.example/docs/COM1.md",
            "https://mdview.example/docs/lpt%C2%B9",
            "https://mdview.example/docs/x.png.",
            "https://mdview.example/docs/x.png%20",
            "https://mdview.example/docs//x.png",
            "https://mdview.example/docs/",
            "https://mdview.example/",
            "https://mdview.example/docs/%00x.png",
            "https://mdview.example/docs/%FF.png",
            "https://unc.mdview.example/server/share",
            "https://unc.mdview.example/./pipe/p",
            "https://unc.mdview.example/%3F/C%3A/x.png",
            "https://unc.mdview.example/server/%5Cshare/x.png",
            "https://drive-dd.mdview.example/x.png",
            "https://drive-1.mdview.example/x.png",
            "https://evil.mdview.example/x.png",
            "https://mdview.example.evil.com/x.png",
            "https://example.com/x.png",
        ] {
            assert_eq!(resolves(url, DOC), None, "{url} must not resolve to a file");
        }
        // Decoded once only: a double-encoded `\` is a literal name on the
        // document's own drive, never a separator.
        assert_eq!(
            resolves("https://mdview.example/%255C%255Cserver%255Cshare%255Cx.png", DOC).as_deref(),
            Some(r"C:\%5C%5Cserver%5Cshare%5Cx.png")
        );
    }

    #[test]
    fn test_paths_that_cannot_be_named_get_no_url() {
        let doc_root = Root::Disk('C');
        for path in [r"\\.\pipe\p", r"\\?\GLOBALROOT\Device\x", r"C:\docs\CON", r"C:\docs\a.png:x", r"C:\docs\x.png.", r"C:\"] {
            assert_eq!(url_for(Path::new(path), &doc_root), None, "{path}");
        }
        // `..` stays below the root.
        assert_eq!(url_for(Path::new(r"D:\a\..\..\..\x.png"), &doc_root).as_deref(), Some("https://drive-d.mdview.example/x.png"));
    }

    #[test]
    fn test_only_the_plain_host_can_be_a_page() {
        assert!(is_document_origin("https://mdview.example/docs/a.md"));
        assert!(is_document_origin("https://MDVIEW.example/docs/a.md"));
        assert!(!is_document_origin("https://drive-d.mdview.example/a.md"));
        assert!(!is_document_origin("https://unc.mdview.example/s/h/a.md"));
        assert!(is_virtual_url("https://unc.mdview.example/s/h/a.md"));
        assert!(!is_virtual_url("https://mdview.example.evil.com/a.md"));
        assert!(!is_virtual_url("https://example.com/a.md"));
    }

    #[test]
    fn test_drive_kinds() {
        // GetDriveTypeW values: 0 unknown, 1 no root, 2 removable, 3 fixed, 4 remote, 5 cd, 6 ram disk.
        assert_eq!(drive_kind_from(3, Some(r"\Device\HarddiskVolume3")), DriveKind::Local);
        assert_eq!(drive_kind_from(2, Some(r"\Device\HarddiskVolume9")), DriveKind::Local);
        assert_eq!(drive_kind_from(5, Some(r"\Device\CdRom0")), DriveKind::Local);
        assert_eq!(drive_kind_from(6, None), DriveKind::Local);
        assert_eq!(drive_kind_from(3, Some(r"\??\C:\Users\me\pics")), DriveKind::Local, "subst to a local folder");
        assert_eq!(drive_kind_from(4, Some(r"\Device\LanmanRedirector\;Z:0000000000012345\server\share")), DriveKind::Network);
        assert_eq!(drive_kind_from(3, Some(r"\??\UNC\server\share\dir")), DriveKind::Network, "drive letter defined on a share");
        assert_eq!(drive_kind_from(3, Some(r"\Device\Mup\server\share")), DriveKind::Network);
        assert_eq!(drive_kind_from(0, None), DriveKind::Network, "unknown counts as network");
        assert_eq!(drive_kind_from(1, None), DriveKind::Missing);
    }

    #[test]
    fn test_location_never_needs_the_network() {
        let probe = |letter: char| match letter {
            'Z' => DriveKind::Network,
            'Q' => DriveKind::Missing,
            _ => DriveKind::Local,
        };
        let doc = Path::new(DOC);
        let at = |p: &str| location(Path::new(p), doc, probe);
        assert_eq!(at(r"C:\x.png"), Some(Location::SameRoot));
        assert_eq!(at(r"c:\x.png"), Some(Location::SameRoot));
        assert_eq!(at(r"D:\x.png"), Some(Location::Local));
        assert_eq!(at(r"Q:\x.png"), Some(Location::Local));
        assert_eq!(at(r"Z:\x.png"), Some(Location::Network("Z:".to_string())));
        assert_eq!(at(r"\\server\share\x.png"), Some(Location::Network(r"\\server\share".to_string())));
        // A document on a share: the same share is its own root, another is network.
        let on_share = Path::new(r"\\server\share\docs\note.md");
        assert_eq!(location(Path::new(r"\\SERVER\share\img\x.png"), on_share, probe), Some(Location::SameRoot));
        assert_eq!(
            location(Path::new(r"\\server\other\x.png"), on_share, probe),
            Some(Location::Network(r"\\server\other".to_string()))
        );
        // The document on a network drive: files on that drive are its own root.
        assert_eq!(location(Path::new(r"Z:\img\x.png"), Path::new(r"Z:\docs\note.md"), probe), Some(Location::SameRoot));
    }

    #[test]
    fn test_absolute_targets_in_documents() {
        for (dest, path) in [
            (r"D:\pictures\x.png", r"D:\pictures\x.png"),
            ("D:/pictures/x.png", r"D:\pictures\x.png"),
            ("file:///D:/my%20pictures/x.png", r"D:\my pictures\x.png"),
            ("FILE:///D:/x.png", r"D:\x.png"),
            ("file://localhost/D:/x.png", r"D:\x.png"),
            (r"\\server\share\x.png", r"\\server\share\x.png"),
            ("file://server/share/x.png", r"\\server\share\x.png"),
        ] {
            assert_eq!(absolute_link_target(dest), Some(PathBuf::from(path)), "{dest}");
        }
        for dest in [
            "img/x.png", "../x.png", "/x.png", "#intro", "https://example.com/x.png", "//example.com/x.png",
            "D:x.png", r"\\?\C:\x.png", r"\\.\pipe\p", "file:////./pipe/p", "file:///%5C%5C.%5Cpipe%5Cp", "mailto:a@b.c",
        ] {
            assert_eq!(absolute_link_target(dest), None, "{dest}");
        }
    }
}
