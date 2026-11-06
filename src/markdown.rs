use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd, html};
use std::collections::HashMap;

pub fn markdown_to_html(markdown: &str) -> String {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);
    options.insert(Options::ENABLE_YAML_STYLE_METADATA_BLOCKS);
    options.insert(Options::ENABLE_PLUSES_DELIMITED_METADATA_BLOCKS);

    let mut events: Vec<Event> = Parser::new_ext(markdown, options).collect();
    assign_heading_ids(&mut events);

    let mut html_output = String::new();
    html::push_html(&mut html_output, events.into_iter());
    html_output
}

/// Give every heading a GitHub-compatible `id`, so that in-document links
/// such as `[Intro](#intro)` written for GitHub work in MDView as well.
fn assign_heading_ids(events: &mut [Event]) {
    let mut occurrences: HashMap<String, usize> = HashMap::new();
    let mut i = 0;

    while i < events.len() {
        if matches!(events[i], Event::Start(Tag::Heading { id: None, .. })) {
            let mut text = String::new();
            let mut j = i + 1;
            while j < events.len() && !matches!(events[j], Event::End(TagEnd::Heading(_))) {
                if let Event::Text(t) | Event::Code(t) = &events[j] {
                    text.push_str(t);
                }
                j += 1;
            }

            let slug = unique_slug(github_slug(&text), &mut occurrences);
            if !slug.is_empty() {
                if let Event::Start(Tag::Heading { id, .. }) = &mut events[i] {
                    *id = Some(slug.into());
                }
            }
            i = j;
        }
        i += 1;
    }
}

/// Slug rules of github-slugger (used by GitHub for heading anchors):
/// lowercase, then keep alphabetic characters, combining marks, decimal
/// digits, connector punctuation (`_`) and `-`, turn each space into `-`
/// (consecutive hyphens are kept) and drop everything else, including
/// "other numbers" such as `½` and `²`.
fn github_slug(text: &str) -> String {
    use unicode_general_category::{GeneralCategory as Gc, get_general_category};

    text.to_lowercase()
        .chars()
        .filter_map(|c| match c {
            ' ' | '-' => Some('-'),
            c if c.is_alphabetic() => Some(c),
            c => match get_general_category(c) {
                Gc::NonspacingMark
                | Gc::SpacingMark
                | Gc::EnclosingMark
                | Gc::DecimalNumber
                | Gc::ConnectorPunctuation => Some(c),
                _ => None,
            },
        })
        .collect()
}

/// De-duplicate slugs the way github-slugger does: `notes`, `notes-1`, `notes-2`, ...
fn unique_slug(base: String, occurrences: &mut HashMap<String, usize>) -> String {
    let mut result = base.clone();
    while occurrences.contains_key(&result) {
        let count = occurrences.entry(base.clone()).or_insert(0);
        *count += 1;
        result = format!("{base}-{count}");
    }
    occurrences.insert(result.clone(), 0);
    result
}

/// Whether a page may load remote (http/https) images.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RemoteImages {
    /// Load them.
    Allowed,
    /// Block them and show a banner offering to load them.
    AskToLoad,
    /// Block them without asking.
    Blocked,
}

/// A bar at the top of a document (Outlook-style), offered by MDView, never by
/// the document. The user's choice is posted as `{type:'bar', bar, choice}`;
/// the host acts on it only if it offered that bar on the page (#40).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bar {
    pub id: &'static str,
    pub text: String,
    pub buttons: Vec<BarButton>,
    pub show: BarShow,
    /// Has a close button ([x], choice `close`). The host keeps a closed bar
    /// away until another document is opened (#24).
    pub closable: bool,
}

/// The answer posted by a bar's close button.
pub const CLOSE_CHOICE: &str = "close";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BarButton {
    pub choice: &'static str,
    pub label: &'static str,
    /// Remove the bar right away (the host does not reload the page).
    pub closes_bar: bool,
}

/// When a bar appears.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarShow {
    /// As soon as the page is shown.
    Now,
    /// Only once a remote image was actually blocked.
    WhenRemoteImageBlocked,
}

pub const REMOTE_IMAGES_BAR: &str = "remote-images";

/// Standalone viewer only: asks to register MDView as .md viewer (#37).
pub const REGISTER_BAR: &str = "register";

/// Mark of the Web (#24): shown instead of a downloaded document until the
/// user chooses to open it.
pub const DOWNLOAD_GATE: &str = "download-gate";

/// What a page needs bars for.
#[derive(Debug, Clone, Copy)]
pub struct PageBars<'a> {
    pub remote_images: RemoteImages,
    /// Only ever true in the standalone viewer.
    pub ask_to_register: bool,
    /// Bars the user closed while this document is open.
    pub closed: &'a [String],
}

/// What the user chose while a document is open (#24): bars they closed, and
/// whether they opened it although it was downloaded from the internet.
/// Reloading the same document keeps these; another document starts fresh.
#[derive(Debug, Default)]
pub struct DocumentChoices {
    document: Option<std::path::PathBuf>,
    closed_bars: Vec<String>,
    opened_download: bool,
}

impl DocumentChoices {
    /// Switch to `document`; choices made for another document are dropped.
    fn switch_to(&mut self, document: &std::path::Path) {
        if self.document.as_deref() != Some(document) {
            self.document = Some(document.to_path_buf());
            self.closed_bars.clear();
            self.opened_download = false;
        }
    }

    /// The closed bars for `document`, which is about to be shown.
    pub fn closed_bars(&mut self, document: &std::path::Path) -> &[String] {
        self.switch_to(document);
        &self.closed_bars
    }

    /// Whether the user chose Open for `document` while it is open.
    pub fn opened_download(&mut self, document: &std::path::Path) -> bool {
        self.switch_to(document);
        self.opened_download
    }

    /// The user closed `bar` on `document`.
    pub fn close_bar(&mut self, document: &std::path::Path, bar: &str) {
        self.switch_to(document);
        if !self.closed_bars.iter().any(|id| id == bar) {
            self.closed_bars.push(bar.to_string());
        }
    }

    /// The user chose Open for the downloaded `document`.
    pub fn open_download(&mut self, document: &std::path::Path) {
        self.switch_to(document);
        self.opened_download = true;
    }
}

/// The gate's question and buttons. None of them close the bar: the host
/// either shows the document or, if that fails, says why and keeps the gate.
pub fn download_gate_bar() -> Bar {
    Bar {
        id: DOWNLOAD_GATE,
        text: "This document was downloaded from the internet.".to_string(),
        buttons: vec![
            BarButton { choice: "open", label: "Open", closes_bar: false },
            BarButton { choice: "unblock", label: "Open and remove the mark", closes_bar: false },
            BarButton { choice: "never", label: "Don't ask again", closes_bar: false },
            BarButton { choice: "closeWindow", label: "Close", closes_bar: false },
        ],
        show: BarShow::Now,
        closable: false,
    }
}

/// The page shown instead of a downloaded document. It contains nothing from
/// the document: only its path and, if known, where it came from, as text.
pub fn download_gate_html(path: &str, host_url: Option<&str>) -> String {
    let mut html = format!("<div class=\"mdview-gate\"><p><code>{}</code></p>", html_text(path));
    if let Some(url) = host_url {
        html.push_str(&format!("<p>From: <code>{}</code></p>", html_text(url)));
    }
    html.push_str("</div>");
    html
}

/// The page shown when a document cannot be read, instead of an empty window.
pub fn read_error_html(path: &str, reason: &str) -> String {
    format!(
        "<div class=\"mdview-gate\"><p><strong>Could not read</strong></p><p><code>{}</code></p><p>{}</p></div>",
        html_text(path),
        html_text(reason)
    )
}

/// Text for HTML element content and attributes.
fn html_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            c => out.push(c),
        }
    }
    out
}

/// The bars MDView puts on a document, top first: registration, remote images.
pub fn document_bars(page: &PageBars) -> Vec<Bar> {
    let mut bars = Vec::new();
    if page.ask_to_register {
        bars.push(Bar {
            id: REGISTER_BAR,
            text: concat!(
                "Do you want to make ",
                crate::settings::app_name!(),
                " the default viewer for Markdown (.md) files?"
            )
            .to_string(),
            buttons: vec![
                BarButton { choice: "register", label: "Register", closes_bar: true },
                BarButton { choice: "notNow", label: "Not now", closes_bar: true },
                BarButton { choice: "never", label: "Don't ask again", closes_bar: true },
            ],
            show: BarShow::Now,
            closable: false,
        });
    }
    if page.remote_images == RemoteImages::AskToLoad {
        bars.push(Bar {
            id: REMOTE_IMAGES_BAR,
            text: "Remote images in this document were blocked to protect your privacy.".to_string(),
            buttons: vec![
                BarButton { choice: "once", label: "Load images", closes_bar: false },
                BarButton { choice: "always", label: "Always load", closes_bar: false },
                BarButton { choice: "never", label: "Never", closes_bar: true },
            ],
            show: BarShow::WhenRemoteImageBlocked,
            closable: true,
        });
    }
    bars.retain(|bar| !page.closed.iter().any(|closed| closed == bar.id));
    bars
}

/// Whether `choice` is an answer to a bar that is on the page.
pub fn bar_choice_offered(bars: &[Bar], bar: &str, choice: &str) -> bool {
    bars.iter().any(|b| {
        b.id == bar
            && ((b.closable && choice == CLOSE_CHOICE)
                || b.buttons.iter().any(|button| button.choice == choice))
    })
}

/// A JSON string literal that is also safe inside an HTML <script>: `<`, `>`
/// and `&` are escaped, so the text can never end the script.
fn script_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '<' | '>' | '&' | '\u{2028}' | '\u{2029}' => out.push_str(&format!("\\u{:04x}", c as u32)),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

fn bars_json(bars: &[Bar]) -> String {
    let items: Vec<String> = bars
        .iter()
        .map(|bar| {
            let buttons: Vec<String> = bar
                .buttons
                .iter()
                .map(|b| {
                    format!(
                        r#"{{"choice":{},"label":{},"closesBar":{}}}"#,
                        script_string(b.choice),
                        script_string(b.label),
                        b.closes_bar
                    )
                })
                .collect();
            let show = match bar.show {
                BarShow::Now => "now",
                BarShow::WhenRemoteImageBlocked => "remoteImageBlocked",
            };
            format!(
                r#"{{"id":{},"text":{},"buttons":[{}],"show":"{show}","closable":{}}}"#,
                script_string(bar.id),
                script_string(&bar.text),
                buttons.join(","),
                bar.closable
            )
        })
        .collect();
    format!("[{}]", items.join(","))
}

/// A random value for the page's Content-Security-Policy nonce. Documents are
/// written before they are shown, so they cannot know it, and only MDView's
/// own scripts carry it. RandomState keys come from the OS random generator.
fn new_nonce() -> String {
    use std::hash::{BuildHasher, Hasher};
    let time = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or_default();
    let part = |salt: u64| {
        let mut h = std::collections::hash_map::RandomState::new().build_hasher();
        h.write_u128(time);
        h.write_u64(salt);
        h.finish()
    };
    format!("{:016x}{:016x}", part(1), part(2))
}

/// Content-Security-Policy for a page: only MDView's own (nonced) scripts run,
/// nothing is fetched or framed, no forms or <base>, and images only come from
/// the document's own files, unless remote images are allowed.
fn content_security_policy(nonce: &str, remote_images: RemoteImages) -> String {
    let img_src = match remote_images {
        RemoteImages::Allowed => "https://mdview.example data: https: http:",
        RemoteImages::AskToLoad | RemoteImages::Blocked => "https://mdview.example data:",
    };
    format!(
        "default-src 'none'; script-src 'nonce-{nonce}'; style-src 'unsafe-inline'; \
         img-src {img_src}; media-src https://mdview.example; font-src 'none'; \
         connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; \
         form-action 'none'"
    )
}

pub fn wrap_html(content: &str, dark_mode: bool, remote_images: RemoteImages, bars: &[Bar]) -> String {
    let nonce = new_nonce();
    let csp = content_security_policy(&nonce, remote_images);
    let bars = bars_json(bars);
    let banner_bg = if dark_mode { "#3b3420" } else { "#fff8c5" };
    let banner_border = if dark_mode { "#6b5b2a" } else { "#d4a72c" };
    let bg_color = if dark_mode { "#1e1e1e" } else { "#ffffff" };
    let text_color = if dark_mode { "#d4d4d4" } else { "#24292e" };
    let code_bg = if dark_mode { "#2d2d2d" } else { "#f6f8fa" };
    let link_color = if dark_mode { "#58a6ff" } else { "#0366d6" };
    let border_color = if dark_mode { "#444" } else { "#e1e4e8" };
    let content = content
        .replace("src=\"/", "src=\"")
        .replace("href=\"/", "href=\"")
        .replace("src='/", "src='")
        .replace("href='/", "href='");

    format!(
        r#"<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="{csp}">
<script nonce="{nonce}">
// Registered before the document's images start loading, so no blocked
// remote image is missed.
window.__mdviewRemoteImageBlocked = false;
document.addEventListener('securitypolicyviolation', function(e) {{
    if (String(e.violatedDirective).indexOf('img-src') === 0 && /^https?:/i.test(e.blockedURI)) {{
        window.__mdviewRemoteImageBlocked = true;
        if (window.__mdviewOnRemoteImageBlocked) window.__mdviewOnRemoteImageBlocked();
    }}
}});
</script>
<style>
#mdview-bars {{
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin: -8px 0 16px 0;
}}
.mdview-bar {{
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
    padding: 6px 10px;
    border: 1px solid {banner_border};
    border-radius: 6px;
    background-color: {banner_bg};
    font-size: 13px;
}}
.mdview-bar span {{ flex: 1; }}
.mdview-bar button {{ font: inherit; cursor: pointer; }}
.mdview-gate code {{ word-break: break-all; }}
.mdview-bar button.mdview-bar-close {{
    border: none;
    background: transparent;
    color: inherit;
    font-size: 16px;
    line-height: 1;
    padding: 0 4px;
}}
body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
    font-size: 14px;
    line-height: 1.6;
    padding: 20px;
    max-width: 900px;
    margin: 0 auto;
    background-color: {bg_color};
    color: {text_color};
}}
a {{ color: {link_color}; text-decoration: none; }}
a:hover {{ text-decoration: underline; }}
code {{
    background-color: {code_bg};
    padding: 0.2em 0.4em;
    border-radius: 3px;
    font-family: "Cascadia Code", "Fira Code", Consolas, monospace;
    font-size: 85%;
}}
pre {{
    background-color: {code_bg};
    padding: 16px;
    overflow: auto;
    border-radius: 6px;
}}
pre code {{
    background-color: transparent;
    padding: 0;
}}
blockquote {{
    border-left: 4px solid {border_color};
    margin: 0;
    padding-left: 16px;
    color: {text_color};
    opacity: 0.8;
}}
table {{
    border-collapse: collapse;
    width: 100%;
}}
th, td {{
    border: 1px solid {border_color};
    padding: 8px 12px;
    text-align: left;
}}
th {{
    background-color: {code_bg};
}}
img {{
    max-width: 100%;
}}
h1, h2 {{
    border-bottom: 1px solid {border_color};
    padding-bottom: 0.3em;
}}
hr {{
    border: none;
    border-top: 1px solid {border_color};
}}
input[type="checkbox"] {{
    margin-right: 0.5em;
}}
a {{ cursor: pointer; }}
</style>
</head>
<body>
{content}
<script nonce="{nonce}">
(function() {{
    // Bars offered by MDView (markdown::Bar), stacked at the top of the page.
    // The answer goes to MDView, which saves it or reloads the page as needed.
    var bars = {bars};
    var webview = window.chrome && window.chrome.webview;

    function showBar(def) {{
        if (!webview || document.querySelector('#mdview-bars > [data-bar="' + def.id + '"]')) return;
        var stack = document.getElementById('mdview-bars');
        if (!stack) {{
            stack = document.createElement('div');
            stack.id = 'mdview-bars';
            document.body.insertBefore(stack, document.body.firstChild);
        }}
        var bar = document.createElement('div');
        bar.className = 'mdview-bar';
        bar.setAttribute('role', 'alert');
        bar.setAttribute('data-bar', def.id);
        var text = document.createElement('span');
        text.textContent = def.text;
        bar.appendChild(text);
        function addButton(label, choice, closesBar, name) {{
            var b = document.createElement('button');
            b.type = 'button';
            b.textContent = label;
            if (name) {{
                b.setAttribute('aria-label', name);
                b.title = name;
                b.className = 'mdview-bar-close';
            }}
            b.addEventListener('click', function() {{
                webview.postMessage({{type: 'bar', bar: def.id, choice: choice}});
                if (closesBar) {{
                    bar.remove();
                    if (!stack.firstChild) stack.remove();
                }}
            }});
            bar.appendChild(b);
        }}
        def.buttons.forEach(function(button) {{
            addButton(button.label, button.choice, button.closesBar);
        }});
        if (def.closable) addButton('×', 'close', true, 'Close');
        stack.appendChild(bar);
    }}

    bars.forEach(function(def) {{ if (def.show === 'now') showBar(def); }});
    window.__mdviewOnRemoteImageBlocked = function() {{
        bars.forEach(function(def) {{ if (def.show === 'remoteImageBlocked') showBar(def); }});
    }};
    if (window.__mdviewRemoteImageBlocked) window.__mdviewOnRemoteImageBlocked();

    // Heading ids are generated server-side (markdown::assign_heading_ids).
    function scrollToHash(hash) {{
        if (!hash || hash.charAt(0) !== '#') return false;

        var id = decodeURIComponent(hash.slice(1));
        if (!id) return false;

        var target = document.getElementById(id);
        if (!target) return false;

        target.scrollIntoView();
        if (history && history.replaceState) {{
            history.replaceState(null, '', hash);
        }}
        return true;
    }}

    if (location.hash) {{
        scrollToHash(location.hash);
    }}

    document.addEventListener('click', function(e) {{
        var link = e.target.closest('a');
        var webview = window.chrome && window.chrome.webview;
        if (!link) return;

        var href = link.getAttribute('href');
        if (!href) return;

        if (href.charAt(0) === '#') {{
            e.preventDefault();
            scrollToHash(href);
            return;
        }}

        if (webview && e.ctrlKey) {{
            e.preventDefault();
            webview.postMessage({{type: 'openLink', url: link.href || href}});
        }}
    }});

    document.addEventListener('keydown', function(e) {{
        var webview = window.chrome && window.chrome.webview;
        if (e.key === 'Escape' && webview) {{
            webview.postMessage({{type: 'close'}});
        }}
    }});
}})();
</script>
</body>
</html>"#
    )
}

#[allow(dead_code)]
pub fn markdown_to_plain_text(markdown: &str) -> String {
    use pulldown_cmark::{Event, Tag, TagEnd};

    let mut options = Options::empty();
    options.insert(Options::ENABLE_YAML_STYLE_METADATA_BLOCKS);
    options.insert(Options::ENABLE_PLUSES_DELIMITED_METADATA_BLOCKS);
    let parser = Parser::new_ext(markdown, options);

    let mut output = String::new();
    let mut in_metadata_block = false;

    for event in parser {
        match event {
            Event::Start(Tag::MetadataBlock(_)) => in_metadata_block = true,
            Event::End(TagEnd::MetadataBlock(_)) => in_metadata_block = false,
            _ if in_metadata_block => {}
            Event::Text(text) => output.push_str(&text),
            Event::Code(code) => {
                output.push('`');
                output.push_str(&code);
                output.push('`');
            }
            Event::SoftBreak | Event::HardBreak => output.push('\n'),
            Event::Start(Tag::Paragraph) => {}
            Event::End(TagEnd::Paragraph) => output.push_str("\n\n"),
            Event::Start(Tag::Heading { .. }) => {}
            Event::End(TagEnd::Heading(_)) => output.push_str("\n\n"),
            Event::Start(Tag::CodeBlock(_)) => output.push_str("\n```\n"),
            Event::End(TagEnd::CodeBlock) => output.push_str("```\n\n"),
            Event::Start(Tag::List(_)) => {}
            Event::End(TagEnd::List(_)) => output.push('\n'),
            Event::Start(Tag::Item) => output.push_str("  - "),
            Event::End(TagEnd::Item) => output.push('\n'),
            Event::Start(Tag::BlockQuote(_)) => output.push_str("> "),
            Event::End(TagEnd::BlockQuote(_)) => output.push('\n'),
            _ => {}
        }
    }

    output.trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn nonce_of(page: &str) -> String {
        let start = page.find("'nonce-").expect("CSP has a nonce") + 7;
        page[start..start + 32].to_string()
    }

    // Security: a document's own <script> and on...= handlers must not run;
    // only MDView's scripts carry the CSP nonce.
    #[test]
    fn test_page_csp_only_lets_mdview_scripts_run() {
        let body = markdown_to_html("# T\n\n<script>evil()</script>\n\n<img src=x onerror=\"evil()\">");
        let page = wrap_html(&body, false, RemoteImages::AskToLoad, &document_bars(&page(RemoteImages::AskToLoad, false)));
        let nonce = nonce_of(&page);
        assert_eq!(nonce.len(), 32);
        let meta_at = page.find("<meta http-equiv=\"Content-Security-Policy\"").expect("CSP meta");
        assert!(meta_at < page.find("<style>").unwrap() && meta_at < page.find("<body>").unwrap());
        for directive in [
            "default-src 'none'",
            &format!("script-src 'nonce-{nonce}';"),
            "connect-src 'none'",
            "frame-src 'none'",
            "object-src 'none'",
            "base-uri 'none'",
            "form-action 'none'",
        ] {
            assert!(page.contains(directive), "missing {directive}");
        }
        assert_eq!(page.matches(&format!("<script nonce=\"{nonce}\">")).count(), 2);
        assert!(page.contains("<script>evil()</script>"), "document script is kept but not nonced");
        assert!(!page.contains("script-src 'unsafe-inline'"));
    }

    // Code that contains HTML or scripts is shown as text, not affected by the
    // CSP; harmless raw HTML such as <details> keeps working.
    #[test]
    fn test_code_with_html_is_shown_as_text() {
        let md = "Inline `<script>alert(1)</script>`\n\n```html\n<script>x()</script>\n<img src=x onerror=\"y()\">\n```\n\n    <script>indented</script>\n\n<details><summary>S</summary>body</details>\n";
        let html = markdown_to_html(md);
        assert!(html.contains("<code>&lt;script&gt;alert(1)&lt;/script&gt;</code>"), "{html}");
        assert!(html.contains("&lt;script&gt;x()&lt;/script&gt;\n&lt;img src=x onerror=\"y()\"&gt;"), "{html}");
        assert!(html.contains("&lt;script&gt;indented&lt;/script&gt;"), "{html}");
        assert!(!html.contains("<script>"), "code must never become a real script tag: {html}");
        assert!(html.contains("<details><summary>S</summary>body</details>"), "{html}");
    }

    #[test]
    fn test_page_nonce_is_new_for_every_page() {
        let a = nonce_of(&wrap_html("", false, RemoteImages::Blocked, &[]));
        let b = nonce_of(&wrap_html("", false, RemoteImages::Blocked, &[]));
        assert_ne!(a, b);
    }

    #[test]
    fn test_remote_images_load_only_when_allowed() {
        let allowed = wrap_html("", false, RemoteImages::Allowed, &[]);
        assert!(allowed.contains("img-src https://mdview.example data: https: http:;"));
        let ask = wrap_html("", false, RemoteImages::AskToLoad, &document_bars(&page(RemoteImages::AskToLoad, false)));
        assert!(ask.contains("img-src https://mdview.example data:;"));
        let blocked = wrap_html("", false, RemoteImages::Blocked, &[]);
        assert!(blocked.contains("img-src https://mdview.example data:;"));
    }

    fn page(remote_images: RemoteImages, ask_to_register: bool) -> PageBars<'static> {
        PageBars { remote_images, ask_to_register, closed: &[] }
    }

    // #40: one bar component; MDView decides which bars a page gets.
    fn bar(id: &'static str, text: &str, buttons: &[(&'static str, &'static str, bool)], show: BarShow) -> Bar {
        Bar {
            id,
            text: text.to_string(),
            buttons: buttons
                .iter()
                .map(|&(choice, label, closes_bar)| BarButton { choice, label, closes_bar })
                .collect(),
            show,
            closable: false,
        }
    }

    #[test]
    fn test_page_shows_every_bar_it_is_given_in_order() {
        let bars = [
            bar("first", "First bar", &[("yes", "Yes", false), ("no", "No", true)], BarShow::Now),
            bar("second", "Second bar", &[("ok", "OK", true)], BarShow::WhenRemoteImageBlocked),
        ];
        let page = wrap_html("", false, RemoteImages::Blocked, &bars);
        let first = page.find(r#""id":"first""#).expect("first bar in the page");
        let second = page.find(r#""id":"second""#).expect("second bar in the page");
        assert!(first < second, "bars keep their order");
        for part in [
            r#""text":"First bar""#,
            r#"{"choice":"yes","label":"Yes","closesBar":false}"#,
            r#"{"choice":"no","label":"No","closesBar":true}"#,
            r#""show":"now""#,
            r#""show":"remoteImageBlocked""#,
        ] {
            assert!(page.contains(part), "missing {part}");
        }
        let none = wrap_html("", false, RemoteImages::Blocked, &[]);
        assert!(none.contains("var bars = [];"), "a page without bars gets an empty list");
    }

    // Bar text is written into MDView's own script; it must never be able to
    // end that script or start another one.
    #[test]
    fn test_bar_text_cannot_leave_the_script() {
        let text = "</script><script>x()</script> & <!-- \"q\" \\ \u{2028}\u{2029}\n";
        let page = wrap_html("", false, RemoteImages::Blocked, &[bar("b", text, &[], BarShow::Now)]);
        let nonce = nonce_of(&page);
        assert_eq!(page.matches("<script").count(), 2, "only MDView's own two scripts");
        assert_eq!(page.matches(&format!("<script nonce=\"{nonce}\">")).count(), 2);
        assert_eq!(page.matches("</script>").count(), 2);
        // Expected JSON, built from parts: `u(x)` is a backslash-u escape.
        let u = |hex: &str| format!("{}u{hex}", char::from(92));
        let bs = char::from(92);
        let expected = format!(
            "\"text\":\"{lt}/script{gt}{lt}script{gt}x(){lt}/script{gt} {amp} {lt}!-- {bs}\"q{bs}\" {bs}{bs} {ls}{ps}{bs}n\"",
            lt = u("003c"), gt = u("003e"), amp = u("0026"), ls = u("2028"), ps = u("2029")
        );
        assert!(page.contains(&expected), "expected {expected}");
    }

    #[test]
    fn test_document_bars_follow_remote_images() {
        let bars = document_bars(&page(RemoteImages::AskToLoad, false));
        assert_eq!(bars.len(), 1);
        let remote = &bars[0];
        assert_eq!(remote.id, REMOTE_IMAGES_BAR);
        assert_eq!(remote.show, BarShow::WhenRemoteImageBlocked);
        assert_eq!(remote.text, "Remote images in this document were blocked to protect your privacy.");
        let buttons: Vec<_> = remote.buttons.iter().map(|b| (b.choice, b.label, b.closes_bar)).collect();
        assert_eq!(
            buttons,
            [("once", "Load images", false), ("always", "Always load", false), ("never", "Never", true)]
        );
        assert!(document_bars(&page(RemoteImages::Allowed, false)).is_empty());
        assert!(document_bars(&page(RemoteImages::Blocked, false)).is_empty());
    }

    // #37: the registration question is a bar on top, not a dialog.
    #[test]
    fn test_registration_bar_comes_first_and_only_when_asked() {
        let bars = document_bars(&page(RemoteImages::AskToLoad, true));
        let ids: Vec<_> = bars.iter().map(|b| b.id).collect();
        assert_eq!(ids, [REGISTER_BAR, REMOTE_IMAGES_BAR], "registration bar on top");
        let register = &bars[0];
        assert_eq!(register.show, BarShow::Now);
        assert_eq!(
            register.text,
            "Do you want to make MDView the default viewer for Markdown (.md) files?"
        );
        let buttons: Vec<_> = register.buttons.iter().map(|b| (b.choice, b.label, b.closes_bar)).collect();
        assert_eq!(
            buttons,
            [("register", "Register", true), ("notNow", "Not now", true), ("never", "Don't ask again", true)]
        );
        assert!(bar_choice_offered(&bars, REGISTER_BAR, "notNow"));
        let without = document_bars(&page(RemoteImages::AskToLoad, false));
        assert!(without.iter().all(|b| b.id != REGISTER_BAR));
        assert!(!bar_choice_offered(&without, REGISTER_BAR, "never"), "not offered, not accepted");
    }

    // The host only acts on an answer to a bar it put on the current page.
    #[test]
    fn test_only_offered_bar_choices_are_accepted() {
        let bars = document_bars(&page(RemoteImages::AskToLoad, false));
        assert!(bar_choice_offered(&bars, REMOTE_IMAGES_BAR, "once"));
        assert!(bar_choice_offered(&bars, REMOTE_IMAGES_BAR, "never"));
        assert!(!bar_choice_offered(&bars, REMOTE_IMAGES_BAR, "delete"), "unknown choice");
        assert!(!bar_choice_offered(&bars, "register", "yes"), "bar not on this page");
        assert!(!bar_choice_offered(&[], REMOTE_IMAGES_BAR, "always"), "no bars offered");
    }

    // #24: a downloaded document is not shown; MDView's own page asks first.
    #[test]
    fn test_download_gate_asks_before_showing_the_document() {
        let gate = download_gate_bar();
        assert_eq!(gate.id, DOWNLOAD_GATE);
        assert_eq!(gate.text, "This document was downloaded from the internet.");
        assert_eq!(gate.show, BarShow::Now);
        assert!(!gate.closable);
        let buttons: Vec<_> = gate.buttons.iter().map(|b| (b.choice, b.label, b.closes_bar)).collect();
        assert_eq!(
            buttons,
            [
                ("open", "Open", false),
                ("unblock", "Open and remove the mark", false),
                ("never", "Don't ask again", false),
                ("closeWindow", "Close", false),
            ],
            "the host decides: it shows the document, or says why not and keeps the gate"
        );
        let bars = [gate];
        for choice in ["open", "unblock", "never", "closeWindow"] {
            assert!(bar_choice_offered(&bars, DOWNLOAD_GATE, choice), "{choice}");
        }
        assert!(!bar_choice_offered(&bars, DOWNLOAD_GATE, "delete"));
        assert!(!bar_choice_offered(&bars, DOWNLOAD_GATE, CLOSE_CHOICE), "the gate has no [x]");
        assert!(!bar_choice_offered(&document_bars(&page(RemoteImages::AskToLoad, true)), DOWNLOAD_GATE, "open"),
            "a document page has no gate, so no answer to it is accepted");
    }

    // #24: the gate shows only the path and the source, as text.
    #[test]
    fn test_download_gate_shows_path_and_source_as_text() {
        let html = download_gate_html(r"C:\dl\<script>x()</script>&.md", Some("https://example.com/a?b=1&c=\"<i>\""));
        assert!(html.contains(r"<code>C:\dl\&lt;script&gt;x()&lt;/script&gt;&amp;.md</code>"), "{html}");
        assert!(html.contains("From: <code>https://example.com/a?b=1&amp;c=&quot;&lt;i&gt;&quot;</code>"), "{html}");
        assert!(!html.contains("<a"), "the source is not a link: {html}");
        assert!(!html.contains("<script"), "{html}");
        let without = download_gate_html(r"C:\dl\a.md", None);
        assert!(!without.contains("From:"), "{without}");
        let page_html = wrap_html(&without, false, RemoteImages::Blocked, &[download_gate_bar()]);
        assert!(page_html.contains(r#""id":"download-gate""#));
    }

    // A document that cannot be read gets a clear page, not an empty window.
    #[test]
    fn test_read_error_page_says_which_file_and_why() {
        let html = read_error_html(r"C:\docs\<b>.md", "denied & <gone>");
        assert!(html.contains("Could not read"), "{html}");
        assert!(html.contains(r"<code>C:\docs\&lt;b&gt;.md</code>"), "{html}");
        assert!(html.contains("<p>denied &amp; &lt;gone&gt;</p>"), "{html}");
    }

    // #24: document bars can be closed; the registration bar has "Not now".
    #[test]
    fn test_only_document_bars_can_be_closed() {
        let bars = document_bars(&page(RemoteImages::AskToLoad, true));
        let closable: Vec<_> = bars.iter().map(|b| (b.id, b.closable)).collect();
        assert_eq!(closable, [(REGISTER_BAR, false), (REMOTE_IMAGES_BAR, true)]);
        assert!(bar_choice_offered(&bars, REMOTE_IMAGES_BAR, CLOSE_CHOICE));
        assert!(!bar_choice_offered(&bars, REGISTER_BAR, CLOSE_CHOICE), "no close button, no close");
        let page_html = wrap_html("", false, RemoteImages::AskToLoad, &bars);
        assert!(page_html.contains(r#""show":"now","closable":false}"#));
        assert!(page_html.contains(r#""show":"remoteImageBlocked","closable":true}"#));
    }

    // #24: choices last while the document is open, and are gone when it is
    // opened again.
    #[test]
    fn test_document_choices_last_until_another_document_is_shown() {
        let a = std::path::Path::new(r"C:\docs\a.md");
        let b = std::path::Path::new(r"C:\docs\b.md");
        let mut choices = DocumentChoices::default();
        assert!(choices.closed_bars(a).is_empty());
        assert!(!choices.opened_download(a));
        choices.close_bar(a, REMOTE_IMAGES_BAR);
        choices.close_bar(a, REMOTE_IMAGES_BAR);
        choices.open_download(a);
        assert_eq!(choices.closed_bars(a), [REMOTE_IMAGES_BAR], "reloading a keeps it closed");
        assert!(choices.opened_download(a), "reloading a keeps it open");
        assert!(choices.closed_bars(b).is_empty(), "another document");
        assert!(!choices.opened_download(a), "a opened again asks again");
        assert!(choices.closed_bars(a).is_empty(), "a opened again shows its bars");
        choices.open_download(b);
        assert!(!choices.opened_download(a), "opening b says nothing about a");
    }

    // #24: a closed bar stays away while its document is open.
    #[test]
    fn test_closed_bars_are_left_out() {
        let closed = [REMOTE_IMAGES_BAR.to_string()];
        let bars = document_bars(&PageBars { closed: &closed, ..page(RemoteImages::AskToLoad, true) });
        let ids: Vec<_> = bars.iter().map(|b| b.id).collect();
        assert_eq!(ids, [REGISTER_BAR]);
        assert!(!bar_choice_offered(&bars, REMOTE_IMAGES_BAR, "once"), "a closed bar takes no answers");
    }

    #[test]
    fn test_basic_markdown() {
        let md = "# Hello\n\nThis is **bold** and *italic*.";
        let html = markdown_to_html(md);
        assert!(html.contains("<h1 id=\"hello\">Hello</h1>"));
        assert!(html.contains("<strong>bold</strong>"));
        assert!(html.contains("<em>italic</em>"));
    }

    #[test]
    fn test_code_block() {
        let md = "```rust\nfn main() {}\n```";
        let html = markdown_to_html(md);
        assert!(html.contains("<code"));
        assert!(html.contains("fn main()"));
    }

    #[test]
    fn test_preserve_relative_urls() {
        let md = "[Guide](../README.md#top)\n\n![Img](../assets/screenshot.png)\n\n[Top](#top)";
        let html = markdown_to_html(md);
        assert!(html.contains("href=\"../README.md#top\""));
        assert!(html.contains("src=\"../assets/screenshot.png\""));
        assert!(html.contains("href=\"#top\""));
    }

    // Issue #4: in-document links written against GitHub-style anchors must
    // match the id generated for the heading.
    #[test]
    fn test_heading_id_matches_github_anchor_issue_4() {
        let md = "1. [STEP 1 — Architecture Analysis](#step-1--architecture-analysis)\n\n\
                  # STEP 1 — Architecture Analysis";
        let html = markdown_to_html(md);
        assert!(html.contains("href=\"#step-1--architecture-analysis\""));
        assert!(
            html.contains("<h1 id=\"step-1--architecture-analysis\">"),
            "heading id does not match link target:\n{html}"
        );
    }

    #[test]
    fn test_heading_id_keeps_non_ascii_letters() {
        let html = markdown_to_html("## Βήμα 1 — Ανάλυση");
        assert!(
            html.contains("<h2 id=\"βήμα-1--ανάλυση\">"),
            "unexpected heading id:\n{html}"
        );
    }

    #[test]
    fn test_heading_id_strips_punctuation_and_includes_inline_code() {
        let html = markdown_to_html("### Run `cargo test`, then *ship*!");
        assert!(
            html.contains("<h3 id=\"run-cargo-test-then-ship\">"),
            "unexpected heading id:\n{html}"
        );
    }

    // Combining marks (vowel signs, viramas) are part of words in many scripts
    // and are kept by GitHub.
    #[test]
    fn test_heading_id_keeps_combining_marks() {
        let html = markdown_to_html("## \u{939}\u{93F}\u{928}\u{94D}\u{926}\u{940} \u{e20}\u{e32}\u{e29}\u{e32}\u{e44}\u{e17}\u{e22}");
        assert!(
            html.contains("<h2 id=\"\u{939}\u{93F}\u{928}\u{94D}\u{926}\u{940}-\u{e20}\u{e32}\u{e29}\u{e32}\u{e44}\u{e17}\u{e22}\">"),
            "unexpected heading id:\n{html}"
        );
    }

    // GitHub drops "other number" characters such as fractions and superscripts.
    #[test]
    fn test_heading_id_drops_other_numbers() {
        let html = markdown_to_html("## \u{BD} half \u{B2} squared");
        assert!(
            html.contains("<h2 id=\"-half--squared\">"),
            "unexpected heading id:\n{html}"
        );
    }

    #[test]
    fn test_duplicate_heading_ids_get_github_suffixes() {
        let html = markdown_to_html("# Notes\n\n# Notes\n\n# Notes");
        assert!(html.contains("<h1 id=\"notes\">"), "{html}");
        assert!(html.contains("<h1 id=\"notes-1\">"), "{html}");
        assert!(html.contains("<h1 id=\"notes-2\">"), "{html}");
    }

    #[test]
    fn test_yaml_front_matter_hidden_from_html() {
        let md = "---\ntags: [rust, markdown]\n---\n\n# Hello";
        let html = markdown_to_html(md);

        assert!(html.contains("<h1 id=\"hello\">Hello</h1>"));
        assert!(!html.contains("tags"));
        assert!(!html.contains("<hr"));
    }

    #[test]
    fn test_toml_front_matter_hidden_from_html() {
        let md = "+++\ntitle = \"My note\"\n+++\n\n# Hello";
        let html = markdown_to_html(md);

        assert!(html.contains("<h1 id=\"hello\">Hello</h1>"));
        assert!(!html.contains("title"), "front matter leaked:\n{html}");
    }

    #[test]
    fn test_plain_text_hides_yaml_front_matter() {
        let md = "---\ntags: [rust, markdown]\n---\n\n# Hello";
        let text = markdown_to_plain_text(md);

        assert!(text.contains("Hello"));
        assert!(!text.contains("tags"));
    }

    #[test]
    fn test_normal_horizontal_rule_still_renders() {
        let md = "# A\n\n---\n\n# B";
        let html = markdown_to_html(md);

        assert!(html.contains("<h1 id=\"a\">A</h1>"));
        assert!(html.contains("<hr"));
        assert!(html.contains("<h1 id=\"b\">B</h1>"));
    }
}
