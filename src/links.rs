//! What a link or a message from the page may make MDView do. Shared by the
//! standalone viewer and the Total Commander plugin.
//!
//! Documents are untrusted: a link must never start a program. Web and mail
//! links go to the default browser/mail app, Markdown files open in MDView,
//! other local files only after the user confirms, and executable or script
//! files are refused outright.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

/// What to do with a link.
#[derive(Debug, PartialEq, Eq)]
pub enum LinkAction {
    /// http, https or mailto: hand to the default app.
    OpenExternal(String),
    /// A Markdown file: show it in MDView.
    OpenMarkdown(PathBuf),
    /// Another local file: open with its default app, but only after the user confirms.
    ConfirmOpenFile(PathBuf),
    /// Never opened; tell the user why.
    Refuse { target: String, reason: &'static str },
}

pub const REFUSED_PROGRAM: &str =
    "MDView does not open programs or scripts from links in a document.";
pub const REFUSED_LINK_TYPE: &str = "MDView does not open this kind of link.";

/// File types that run code when opened. Compared case-insensitively against
/// the extension after `normalized_file_name`.
const EXECUTABLE_EXTENSIONS: &[&str] = &[
    // PATHEXT and other directly executable types
    "exe", "com", "bat", "cmd", "pif", "scr", "cpl", "msc", "msi", "msp", "mst", "appx",
    "appxbundle", "msix", "msixbundle", "application", "appref-ms", "gadget", "jar",
    // scripts
    "vbs", "vbe", "js", "jse", "wsf", "wsh", "wsc", "ws", "ps1", "psm1", "psd1", "ps1xml",
    "psc1", "msh", "msh1", "msh2", "mshxml", "sct", "hta", "py", "pyw",
    // files that run or redirect something when opened
    "lnk", "url", "website", "scf", "reg", "inf", "ins", "isp", "settingcontent-ms",
    "library-ms", "search-ms", "searchconnector-ms", "diagcab", "diagcfg", "diagpkg", "chm",
    "hlp", "iqy", "slk", "xll", "xbap", "job", "dll", "ocx", "sys", "cab",
];

/// The file name as Windows will actually open it: without an alternate data
/// stream suffix (`a.exe:x`, `a.exe::$DATA`) and without trailing dots and
/// spaces, which Windows ignores (`a.exe.` and `a.exe ` open `a.exe`).
fn normalized_file_name(path: &Path) -> String {
    let name = path
        .file_name()
        .map(|n| n.to_string_lossy().into_owned())
        .unwrap_or_default();
    let name = name.split(':').next().unwrap_or_default();
    name.trim_end_matches(['.', ' ']).to_string()
}

fn extension_of(path: &Path) -> String {
    let name = normalized_file_name(path);
    match name.rsplit_once('.') {
        Some((stem, ext)) if !stem.is_empty() => ext.to_ascii_lowercase(),
        _ => String::new(),
    }
}

pub fn is_markdown(path: &Path) -> bool {
    matches!(extension_of(path).as_str(), "md" | "markdown")
}

pub fn is_executable(path: &Path) -> bool {
    EXECUTABLE_EXTENSIONS.contains(&extension_of(path).as_str())
}

/// `a.txt:evil.exe` / `a.exe::$DATA`: an NTFS alternate data stream. No real
/// document links to one, and Windows may judge it by the stream's extension.
fn names_data_stream(path: &Path) -> bool {
    path.file_name()
        .is_some_and(|n| n.to_string_lossy().contains(':'))
}

/// Decide what to do with a local file a link points to.
pub fn classify_local_file(path: &Path) -> LinkAction {
    if names_data_stream(path) {
        LinkAction::Refuse {
            target: path.display().to_string(),
            reason: REFUSED_LINK_TYPE,
        }
    } else if is_executable(path) {
        LinkAction::Refuse {
            target: path.display().to_string(),
            reason: REFUSED_PROGRAM,
        }
    } else if is_markdown(path) {
        LinkAction::OpenMarkdown(path.to_path_buf())
    } else {
        LinkAction::ConfirmOpenFile(path.to_path_buf())
    }
}

/// Decide what to do with a link URL from the page. `resolve_virtual` maps
/// MDView's own `https://mdview.example/...` URLs to the local file they stand
/// for (`None` for any other URL).
pub fn classify_url(url: &str, resolve_virtual: impl Fn(&str) -> Option<PathBuf>) -> LinkAction {
    if let Some(path) = resolve_virtual(url) {
        return classify_local_file(&path);
    }
    let lower = url.trim_start().to_ascii_lowercase();
    if lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:") {
        return LinkAction::OpenExternal(url.trim().to_string());
    }
    if lower.starts_with("file:") {
        if let Some(path) = file_url_to_path(url) {
            return classify_local_file(&path);
        }
    }
    LinkAction::Refuse {
        target: url.to_string(),
        reason: REFUSED_LINK_TYPE,
    }
}

/// `file:///C:/dir/a%20b.txt` or `file://server/share/x` to a Windows path.
fn file_url_to_path(url: &str) -> Option<PathBuf> {
    let rest = url.get(5..)?; // after "file:"
    let rest = rest.split(['?', '#']).next().unwrap_or_default();
    let decoded = percent_decode(rest);
    let path = if let Some(local) = decoded.strip_prefix("///") {
        local.replace('/', "\\")
    } else if let Some(unc) = decoded.strip_prefix("//") {
        format!("\\\\{}", unc.replace('/', "\\"))
    } else {
        return None;
    };
    (!path.is_empty()).then(|| PathBuf::from(path))
}

fn percent_decode(input: &str) -> String {
    let bytes = input.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let (Some(h), Some(l)) = (hex(bytes[i + 1]), hex(bytes[i + 2])) {
                out.push(h << 4 | l);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn hex(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}

/// Carry out a link decision. `owner` is the window dialogs belong to;
/// `open_markdown` shows a Markdown file in the viewer. Every outcome other
/// than success is reported to the user.
pub fn perform(
    action: LinkAction,
    owner: windows::Win32::Foundation::HWND,
    open_markdown: impl FnOnce(&Path),
) {
    use windows::Win32::UI::WindowsAndMessaging::{
        IDYES, MB_DEFBUTTON2, MB_ICONWARNING, MB_OK, MB_YESNO,
    };
    match action {
        LinkAction::OpenExternal(url) => shell_open(owner, &url),
        LinkAction::OpenMarkdown(path) => open_markdown(&path),
        LinkAction::ConfirmOpenFile(path) => {
            let text = format!(
                "This link opens a file on your computer with its default program:\n\n{}\n\nOnly open it if you trust this document. Open it?",
                path.display()
            );
            if message_box(owner, &text, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) == IDYES {
                shell_open(owner, &path.to_string_lossy());
            }
        }
        LinkAction::Refuse { target, reason } => {
            message_box(owner, &format!("{reason}\n\n{target}"), MB_OK | MB_ICONWARNING);
        }
    }
}

fn message_box(
    owner: windows::Win32::Foundation::HWND,
    text: &str,
    style: windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_STYLE,
) -> windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_RESULT {
    use windows::Win32::UI::WindowsAndMessaging::MessageBoxW;
    use windows::core::PCWSTR;
    let text: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let title: Vec<u16> = "MDView\0".encode_utf16().collect();
    unsafe { MessageBoxW(Some(owner), PCWSTR(text.as_ptr()), PCWSTR(title.as_ptr()), style) }
}

fn shell_open(owner: windows::Win32::Foundation::HWND, target: &str) {
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::{MB_ICONERROR, MB_OK, SW_SHOWNORMAL};
    use windows::core::PCWSTR;
    let target_wide: Vec<u16> = target.encode_utf16().chain(std::iter::once(0)).collect();
    let verb: Vec<u16> = "open\0".encode_utf16().collect();
    let result = unsafe {
        ShellExecuteW(
            Some(owner),
            PCWSTR(verb.as_ptr()),
            PCWSTR(target_wide.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        )
    };
    // ShellExecute returns a value greater than 32 on success.
    if result.0 as usize <= 32 {
        message_box(owner, &format!("Could not open:\n\n{target}"), MB_OK | MB_ICONERROR);
    }
}

/// A message posted by MDView's page script: a flat JSON object of strings,
/// e.g. `{"type":"openLink","url":"https://example.com"}`. Anything else
/// (other value types, nesting, invalid JSON) is rejected.
pub fn parse_page_message(json: &str) -> Option<HashMap<String, String>> {
    let mut p = JsonReader { s: json.as_bytes(), i: 0 };
    p.ws();
    p.expect(b'{')?;
    let mut map = HashMap::new();
    p.ws();
    if p.peek() == Some(b'}') {
        p.i += 1;
    } else {
        loop {
            p.ws();
            let key = p.string()?;
            p.ws();
            p.expect(b':')?;
            p.ws();
            let value = p.string()?;
            map.insert(key, value);
            p.ws();
            match p.next()? {
                b',' => continue,
                b'}' => break,
                _ => return None,
            }
        }
    }
    p.ws();
    (p.i == p.s.len()).then_some(map)
}

struct JsonReader<'a> {
    s: &'a [u8],
    i: usize,
}

impl JsonReader<'_> {
    fn peek(&self) -> Option<u8> {
        self.s.get(self.i).copied()
    }
    fn next(&mut self) -> Option<u8> {
        let b = self.peek()?;
        self.i += 1;
        Some(b)
    }
    fn expect(&mut self, b: u8) -> Option<()> {
        (self.next()? == b).then_some(())
    }
    fn ws(&mut self) {
        while matches!(self.peek(), Some(b' ' | b'\t' | b'\n' | b'\r')) {
            self.i += 1;
        }
    }
    fn hex4(&mut self) -> Option<u32> {
        let mut v = 0;
        for _ in 0..4 {
            v = v << 4 | u32::from(hex(self.next()?)?);
        }
        Some(v)
    }
    fn string(&mut self) -> Option<String> {
        self.expect(b'"')?;
        let mut out = String::new();
        loop {
            let b = self.next()?;
            match b {
                b'"' => return Some(out),
                b'\\' => match self.next()? {
                    b'"' => out.push('"'),
                    b'\\' => out.push('\\'),
                    b'/' => out.push('/'),
                    b'b' => out.push('\u{8}'),
                    b'f' => out.push('\u{c}'),
                    b'n' => out.push('\n'),
                    b'r' => out.push('\r'),
                    b't' => out.push('\t'),
                    b'u' => {
                        let hi = self.hex4()?;
                        let c = if (0xD800..0xDC00).contains(&hi) {
                            self.expect(b'\\')?;
                            self.expect(b'u')?;
                            let lo = self.hex4()?;
                            if !(0xDC00..0xE000).contains(&lo) {
                                return None;
                            }
                            char::from_u32(0x10000 + ((hi - 0xD800) << 10) + (lo - 0xDC00))?
                        } else {
                            char::from_u32(hi)?
                        };
                        out.push(c);
                    }
                    _ => return None,
                },
                0x00..=0x1F => return None,
                _ => {
                    // copy one UTF-8 sequence
                    let start = self.i - 1;
                    let len = match b {
                        0x00..=0x7F => 1,
                        0xC0..=0xDF => 2,
                        0xE0..=0xEF => 3,
                        _ => 4,
                    };
                    let end = start + len;
                    let piece = std::str::from_utf8(self.s.get(start..end)?).ok()?;
                    out.push_str(piece);
                    self.i = end;
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn no_virtual(_: &str) -> Option<PathBuf> {
        None
    }

    fn refused(a: &LinkAction) -> bool {
        matches!(a, LinkAction::Refuse { .. })
    }

    #[test]
    fn test_web_and_mail_links_open_externally() {
        for url in ["https://example.com/a", "http://example.com", "mailto:a@b.c", "HTTPS://X.Y"] {
            assert!(matches!(classify_url(url, no_virtual), LinkAction::OpenExternal(_)), "{url}");
        }
    }

    #[test]
    fn test_other_schemes_and_bare_paths_are_refused() {
        for url in [
            "calculator:",
            "ms-msdt:/id x",
            "search-ms:query=x",
            "javascript:alert(1)",
            "data:text/html,x",
            "C:\\Windows\\System32\\calc.exe",
            "\\\\server\\share\\x.exe",
            "shell:startup",
            "",
        ] {
            assert!(refused(&classify_url(url, no_virtual)), "{url:?} must be refused");
        }
    }

    #[test]
    fn test_file_urls_follow_the_local_file_rules() {
        assert!(refused(&classify_url("file:///C:/tmp/canary.cmd", no_virtual)));
        assert!(refused(&classify_url("file://server/share/x.exe", no_virtual)));
        assert_eq!(
            classify_url("file:///C:/tmp/a%20b.pdf", no_virtual),
            LinkAction::ConfirmOpenFile(PathBuf::from("C:\\tmp\\a b.pdf"))
        );
        assert_eq!(
            classify_url("file:///C:/tmp/notes.md", no_virtual),
            LinkAction::OpenMarkdown(PathBuf::from("C:\\tmp\\notes.md"))
        );
    }

    #[test]
    fn test_virtual_urls_use_the_resolved_file() {
        let resolve = |u: &str| u.strip_prefix("https://mdview.example/").map(|p| PathBuf::from(format!("C:\\{p}")));
        assert!(refused(&classify_url("https://mdview.example/tmp/canary.cmd", resolve)));
        assert_eq!(
            classify_url("https://mdview.example/tmp/next.md", resolve),
            LinkAction::OpenMarkdown(PathBuf::from("C:\\tmp/next.md"))
        );
        assert!(matches!(
            classify_url("https://mdview.example/tmp/report.pdf", resolve),
            LinkAction::ConfirmOpenFile(_)
        ));
    }

    #[test]
    fn test_executables_are_refused_in_any_disguise() {
        for name in [
            "a.exe", "A.EXE", "a.cmd", "a.bat", "a.ps1", "a.lnk", "a.url", "a.hta", "a.vbs",
            "a.js", "a.msi", "a.scr", "a.cpl", "a.reg", "a.jar", "a.appref-ms",
            "a.settingcontent-ms", "a.library-ms", "a.search-ms", "a.pdf.exe",
            // Windows ignores trailing dots/spaces and opens the stream's file
            "a.exe.", "a.exe ", "a.exe. . ", "a.exe:stream", "a.exe::$DATA", "a.txt:evil.exe",
        ] {
            let path = PathBuf::from(format!("C:\\tmp\\{name}"));
            assert!(refused(&classify_local_file(&path)), "{name:?} must be refused");
        }
    }

    #[test]
    fn test_documents_need_confirmation_and_markdown_opens() {
        for name in ["a.pdf", "a.png", "a.txt", "a.csv", "a.html", "noextension", "a.MD.txt"] {
            let path = PathBuf::from(format!("C:\\tmp\\{name}"));
            assert!(matches!(classify_local_file(&path), LinkAction::ConfirmOpenFile(_)), "{name}");
        }
        for name in ["a.md", "a.MD", "a.markdown"] {
            let path = PathBuf::from(format!("C:\\tmp\\{name}"));
            assert!(matches!(classify_local_file(&path), LinkAction::OpenMarkdown(_)), "{name}");
        }
    }

    #[test]
    fn test_page_messages_are_parsed_exactly() {
        let m = parse_page_message(r#"{"type":"openLink","url":"https://x/\"close\"?a=\u00e9&b=\ud83d\ude80"}"#).unwrap();
        assert_eq!(m["type"], "openLink");
        assert_eq!(m["url"], "https://x/\"close\"?a=é&b=🚀");
        // a value of "close" in another field is not a close message
        let m = parse_page_message(r#"{"type":"note","text":"close"}"#).unwrap();
        assert_eq!(m["type"], "note");
        assert_eq!(parse_page_message(r#"{}"#).unwrap().len(), 0);
    }

    #[test]
    fn test_invalid_page_messages_are_rejected() {
        for json in [
            "",
            "\"close\"",
            r#"{"type":"close""#,
            r#"{"type":1}"#,
            r#"{"type":{"x":"y"}}"#,
            r#"{"type":"a"} trailing"#,
            r#"{"type":"bad \x escape"}"#,
            "{\"type\":\"raw\u{1}control\"}",
        ] {
            assert!(parse_page_message(json).is_none(), "{json:?} must be rejected");
        }
    }
}
