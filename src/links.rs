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
    /// A Markdown or other file on another share or network drive: opened
    /// (Markdown in MDView, other files with their default app) only after
    /// the user confirms, because opening it connects to that location.
    ConfirmOpenNetwork { path: PathBuf, location: String },
    /// A program or script. A plain click never runs it; Ctrl+click runs it
    /// after a clear warning (1.0.0). `location` is set when it is on another
    /// share or a network drive.
    Program { path: PathBuf, location: Option<String> },
    /// Never opened; tell the user why.
    Refuse { target: String, reason: &'static str },
}

/// What the user agreed to in a question about a link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Confirmed {
    /// Open with its default program (a local file or one on a network location).
    OpenFile(PathBuf),
    /// Show a Markdown file from a network location in MDView.
    OpenMarkdown(PathBuf),
    /// Start a program the user Ctrl+clicked.
    RunProgram(PathBuf),
}

use crate::messages::{Message, Tone};
use crate::settings::app_name;

pub const REFUSED_PROGRAM: &str =
    concat!(app_name!(), " does not open programs or scripts from links in a document.");
pub const REFUSED_LINK_TYPE: &str = concat!(app_name!(), " does not open this kind of link.");

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

/// Decide what to do with a file a link points to. `network` names the
/// network location it is on (another share or a network drive), if any.
pub fn classify_local_file(path: &Path, network: Option<String>) -> LinkAction {
    if names_data_stream(path) {
        LinkAction::Refuse {
            target: path.display().to_string(),
            reason: REFUSED_LINK_TYPE,
        }
    } else if is_executable(path) {
        LinkAction::Program { path: path.to_path_buf(), location: network }
    } else if let Some(location) = network {
        LinkAction::ConfirmOpenNetwork { path: path.to_path_buf(), location }
    } else if is_markdown(path) {
        LinkAction::OpenMarkdown(path.to_path_buf())
    } else {
        LinkAction::ConfirmOpenFile(path.to_path_buf())
    }
}

/// Decide what to do with a link URL from the page. `resolve_virtual` maps
/// MDView's own `https://mdview.example/...` URLs to the local file they stand
/// for (`None` for any other URL).
pub fn classify_url(
    url: &str,
    resolve_virtual: impl Fn(&str) -> Option<PathBuf>,
    network_location: impl Fn(&Path) -> Option<String>,
) -> LinkAction {
    let refuse = || LinkAction::Refuse {
        target: url.to_string(),
        reason: REFUSED_LINK_TYPE,
    };
    // One of MDView's own URLs that names no plain file is refused, never
    // handed to the browser as a web link.
    if crate::vpath::is_virtual_url(url) {
        return match resolve_virtual(url) {
            Some(path) => classify_local_file(&path, network_location(&path)),
            None => refuse(),
        };
    }
    let lower = url.trim_start().to_ascii_lowercase();
    if lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:") {
        return LinkAction::OpenExternal(url.trim().to_string());
    }
    if lower.starts_with("file:") {
        if let Some(path) = crate::vpath::file_url_to_path(url) {
            return classify_local_file(&path, network_location(&path));
        }
    }
    refuse()
}

fn hex(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}
/// The question or message for a link decision, or `None` when there is
/// nothing to ask (web links, Markdown files). `ctrl`: the user Ctrl+clicked.
pub fn question(action: &LinkAction, ctrl: bool) -> Option<(Message, Option<Confirmed>)> {
    match action {
        LinkAction::OpenExternal(_) | LinkAction::OpenMarkdown(_) => None,
        LinkAction::ConfirmOpenFile(path) => Some((
            Message::confirm(
                Tone::Warning,
                format!(
                    "This link opens a file on your computer with its default program:\n\n{}\n\nOnly open it if you trust this document.",
                    path.display()
                ),
                "Open",
            ),
            Some(Confirmed::OpenFile(path.clone())),
        )),
        LinkAction::ConfirmOpenNetwork { path, location } => {
            let (text, confirmed) = if is_markdown(path) {
                (
                    format!(
                        "This link opens a document from a network location:\n\n{}\n\nOpening it connects to {location}. Only open it if you trust this document.",
                        path.display()
                    ),
                    Confirmed::OpenMarkdown(path.clone()),
                )
            } else {
                (
                    format!(
                        "This link opens a file from a network location with its default program:\n\n{}\n\nOpening it connects to {location}. Only open it if you trust this document.",
                        path.display()
                    ),
                    Confirmed::OpenFile(path.clone()),
                )
            };
            Some((Message::confirm(Tone::Warning, text, "Open"), Some(confirmed)))
        }
        LinkAction::Program { path, location } if ctrl => {
            let network = location
                .as_ref()
                .map(|l| format!("\n\nRunning it connects to {l}."))
                .unwrap_or_default();
            Some((
                Message::confirm(
                    Tone::Error,
                    format!(
                        "This link starts a program:\n\n{}\n\nPrograms from a document can harm your computer. Only run it if you trust this document and the program.{network}",
                        path.display()
                    ),
                    "Run program",
                ),
                Some(Confirmed::RunProgram(path.clone())),
            ))
        }
        LinkAction::Program { path, .. } => Some((
            Message::warning(format!(
                "{REFUSED_PROGRAM}\n\n{}\n\nTo run it anyway, Ctrl+click the link.",
                path.display()
            )),
            None,
        )),
        LinkAction::Refuse { target, reason } => Some((Message::warning(format!("{reason}\n\n{target}")), None)),
    }
}

/// Carry out a link decision. `open_markdown` shows a Markdown file in the
/// viewer; `tell` shows a message or asks a question (in the page), and a
/// "yes" comes back through `run_confirmed`. Every outcome other than
/// success is reported to the user.
pub fn perform(
    action: LinkAction,
    ctrl: bool,
    owner: windows::Win32::Foundation::HWND,
    open_markdown: impl FnOnce(&Path),
    tell: impl FnOnce(Message, Option<Confirmed>),
) {
    if let Some((message, confirmed)) = question(&action, ctrl) {
        return tell(message, confirmed);
    }
    match action {
        LinkAction::OpenExternal(url) => {
            if let Err(message) = shell_open(owner, &url) {
                tell(message, None);
            }
        }
        LinkAction::OpenMarkdown(path) => open_markdown(&path),
        _ => {}
    }
}

/// The user said yes to a question from `perform`.
pub fn run_confirmed(
    confirmed: Confirmed,
    owner: windows::Win32::Foundation::HWND,
    open_markdown: impl FnOnce(&Path),
    tell: impl FnOnce(Message, Option<Confirmed>),
) {
    let target = match confirmed {
        Confirmed::OpenMarkdown(path) => return open_markdown(&path),
        Confirmed::OpenFile(path) | Confirmed::RunProgram(path) => path,
    };
    if let Err(message) = shell_open(owner, &target.to_string_lossy()) {
        tell(message, None);
    }
}

/// Open with the default program, as Explorer would: Windows' own checks
/// (SmartScreen, Mark of the Web) apply.
fn shell_open(owner: windows::Win32::Foundation::HWND, target: &str) -> Result<(), Message> {
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;
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
        Err(Message::error(format!("Could not open:\n\n{target}")))
    } else {
        Ok(())
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

    fn not_network(_: &Path) -> Option<String> {
        None
    }

    /// Not opened by a plain click: refused, or a program (Ctrl+click only).
    fn refused(a: &LinkAction) -> bool {
        matches!(a, LinkAction::Refuse { .. } | LinkAction::Program { .. })
    }

    // 1.0.0: a program link is never run by a plain click; Ctrl+click asks
    // with a clear warning, Cancel being the safe answer. Data streams and
    // non-file links are never run, not even with Ctrl+click.
    #[test]
    fn test_programs_run_only_after_ctrl_click_and_a_warning() {
        let program = classify_local_file(Path::new(r"C:\tools\setup.exe"), None);
        assert_eq!(program, LinkAction::Program { path: PathBuf::from(r"C:\tools\setup.exe"), location: None });
        let (plain, action) = question(&program, false).unwrap();
        assert!(action.is_none() && !plain.is_question());
        assert!(plain.text.contains("Ctrl+click") && plain.text.contains(r"C:\tools\setup.exe"), "{}", plain.text);
        let (warning, action) = question(&program, true).unwrap();
        assert!(warning.is_question() && warning.safe == crate::messages::NO);
        assert!(warning.text.contains("can harm your computer") && warning.buttons[0].label == "Run program", "{warning:?}");
        assert_eq!(action, Some(Confirmed::RunProgram(PathBuf::from(r"C:\tools\setup.exe"))));
        let network = classify_local_file(Path::new(r"\\server\share\x.cmd"), Some(r"\\server\share".into()));
        assert!(question(&network, true).unwrap().0.text.contains(r"connects to \\server\share"));
        for never in [r"C:\tmp\a.exe:stream", r"C:\tmp\a.txt:evil.exe", r"C:\tmp\a.exe::$DATA"] {
            let a = classify_local_file(Path::new(never), None);
            assert!(matches!(a, LinkAction::Refuse { .. }), "{never}");
            assert!(question(&a, true).unwrap().1.is_none(), "{never}: nothing to run with Ctrl+click");
        }
        let scheme = classify_url("ms-msdt:/id x", no_virtual, not_network);
        assert!(question(&scheme, true).unwrap().1.is_none());
    }

    #[test]
    fn test_questions_for_files_and_network_locations() {
        let file = LinkAction::ConfirmOpenFile(PathBuf::from(r"C:\a.pdf"));
        let (m, a) = question(&file, false).unwrap();
        assert!(m.is_question() && m.safe == crate::messages::NO && a == Some(Confirmed::OpenFile(PathBuf::from(r"C:\a.pdf"))));
        let md = LinkAction::ConfirmOpenNetwork { path: PathBuf::from(r"\\s\h\a.md"), location: r"\\s\h".into() };
        assert_eq!(question(&md, false).unwrap().1, Some(Confirmed::OpenMarkdown(PathBuf::from(r"\\s\h\a.md"))));
        assert!(question(&LinkAction::OpenMarkdown(PathBuf::from("a.md")), true).is_none());
        assert!(question(&LinkAction::OpenExternal("https://x".into()), false).is_none());
    }

    #[test]
    fn test_web_and_mail_links_open_externally() {
        for url in ["https://example.com/a", "http://example.com", "mailto:a@b.c", "HTTPS://X.Y"] {
            assert!(matches!(classify_url(url, no_virtual, not_network), LinkAction::OpenExternal(_)), "{url}");
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
            assert!(refused(&classify_url(url, no_virtual, not_network)), "{url:?} must be refused");
        }
    }

    #[test]
    fn test_file_urls_follow_the_local_file_rules() {
        assert!(refused(&classify_url("file:///C:/tmp/canary.cmd", no_virtual, not_network)));
        assert!(refused(&classify_url("file://server/share/x.exe", no_virtual, not_network)));
        assert_eq!(
            classify_url("file:///C:/tmp/a%20b.pdf", no_virtual, not_network),
            LinkAction::ConfirmOpenFile(PathBuf::from("C:\\tmp\\a b.pdf"))
        );
        assert_eq!(
            classify_url("file:///C:/tmp/notes.md", no_virtual, not_network),
            LinkAction::OpenMarkdown(PathBuf::from("C:\\tmp\\notes.md"))
        );
    }

    #[test]
    fn test_virtual_urls_use_the_resolved_file() {
        let resolve = |u: &str| u.strip_prefix("https://mdview.example/").map(|p| PathBuf::from(format!("C:\\{p}")));
        assert!(refused(&classify_url("https://mdview.example/tmp/canary.cmd", resolve, not_network)));
        assert_eq!(
            classify_url("https://mdview.example/tmp/next.md", resolve, not_network),
            LinkAction::OpenMarkdown(PathBuf::from("C:\\tmp/next.md"))
        );
        assert!(matches!(
            classify_url("https://mdview.example/tmp/report.pdf", resolve, not_network),
            LinkAction::ConfirmOpenFile(_)
        ));
    }

    // A refused MDView URL is never handed to the browser as a web link.
    #[test]
    fn test_unresolvable_virtual_urls_are_refused() {
        for url in [
            "https://mdview.example/%5C%5Cserver%5Cshare%5Cx.md",
            "https://drive-d.mdview.example/CON",
            "https://unc.mdview.example/server",
        ] {
            assert!(refused(&classify_url(url, no_virtual, not_network)), "{url}");
        }
    }

    // Opening anything on another share or network drive connects to it, so
    // the user is asked first; programs stay refused.
    #[test]
    fn test_network_locations_need_confirmation() {
        let share = |p: &Path| p.to_string_lossy().starts_with(r"\\").then(|| r"\\server\share".to_string());
        for url in ["file://server/share/notes.md", "file://server/share/report.pdf"] {
            assert!(matches!(
                classify_url(url, no_virtual, share),
                LinkAction::ConfirmOpenNetwork { ref location, .. } if location == r"\\server\share"
            ), "{url}");
        }
        assert!(refused(&classify_url("file://server/share/x.exe", no_virtual, share)));
        assert!(refused(&classify_url("file://server/share/a.md:x", no_virtual, share)));
        assert_eq!(
            classify_url("file:///C:/tmp/notes.md", no_virtual, share),
            LinkAction::OpenMarkdown(PathBuf::from(r"C:\tmp\notes.md"))
        );
        // Special paths are not links to a file at all.
        assert!(refused(&classify_url("file:////./pipe/p", no_virtual, share)));
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
            assert!(refused(&classify_local_file(&path, None)), "{name:?} must be refused");
        }
    }

    #[test]
    fn test_documents_need_confirmation_and_markdown_opens() {
        for name in ["a.pdf", "a.png", "a.txt", "a.csv", "a.html", "noextension", "a.MD.txt"] {
            let path = PathBuf::from(format!("C:\\tmp\\{name}"));
            assert!(matches!(classify_local_file(&path, None), LinkAction::ConfirmOpenFile(_)), "{name}");
        }
        for name in ["a.md", "a.MD", "a.markdown"] {
            let path = PathBuf::from(format!("C:\\tmp\\{name}"));
            assert!(matches!(classify_local_file(&path, None), LinkAction::OpenMarkdown(_)), "{name}");
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
