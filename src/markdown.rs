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

pub fn wrap_html(content: &str, dark_mode: bool, remote_images: RemoteImages) -> String {
    let nonce = new_nonce();
    let csp = content_security_policy(&nonce, remote_images);
    let ask_remote_images = remote_images == RemoteImages::AskToLoad;
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
window.__mdviewBlockedRemoteImages = false;
document.addEventListener('securitypolicyviolation', function(e) {{
    if (String(e.violatedDirective).indexOf('img-src') === 0 && /^https?:/i.test(e.blockedURI)) {{
        window.__mdviewBlockedRemoteImages = true;
        if (window.__mdviewShowRemoteImagesBanner) window.__mdviewShowRemoteImagesBanner();
    }}
}});
</script>
<style>
#mdview-remote-images {{
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 8px;
    margin: -8px 0 16px 0;
    padding: 6px 10px;
    border: 1px solid {banner_border};
    border-radius: 6px;
    background-color: {banner_bg};
    font-size: 13px;
}}
#mdview-remote-images span {{ flex: 1; }}
#mdview-remote-images button {{ font: inherit; cursor: pointer; }}
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
    var askRemoteImages = {ask_remote_images};

    // Outlook-style bar when remote images were blocked: load them this time,
    // always, or never ask again. MDView itself reloads the page as needed.
    window.__mdviewShowRemoteImagesBanner = function() {{
        var webview = window.chrome && window.chrome.webview;
        if (!askRemoteImages || !webview || document.getElementById('mdview-remote-images')) return;
        var bar = document.createElement('div');
        bar.id = 'mdview-remote-images';
        bar.setAttribute('role', 'alert');
        var text = document.createElement('span');
        text.textContent = 'Remote images in this document were blocked to protect your privacy.';
        bar.appendChild(text);
        [['once', 'Load images'], ['always', 'Always load'], ['never', 'Never']].forEach(function(c) {{
            var b = document.createElement('button');
            b.type = 'button';
            b.textContent = c[1];
            b.addEventListener('click', function() {{
                webview.postMessage({{type: 'remoteImages', choice: c[0]}});
                if (c[0] === 'never') bar.remove();
            }});
            bar.appendChild(b);
        }});
        document.body.insertBefore(bar, document.body.firstChild);
    }};
    if (window.__mdviewBlockedRemoteImages) window.__mdviewShowRemoteImagesBanner();

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
        let page = wrap_html(&body, false, RemoteImages::AskToLoad);
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
        let a = nonce_of(&wrap_html("", false, RemoteImages::Blocked));
        let b = nonce_of(&wrap_html("", false, RemoteImages::Blocked));
        assert_ne!(a, b);
    }

    #[test]
    fn test_remote_images_load_only_when_allowed() {
        let allowed = wrap_html("", false, RemoteImages::Allowed);
        assert!(allowed.contains("img-src https://mdview.example data: https: http:;"));
        assert!(allowed.contains("var askRemoteImages = false"));
        let ask = wrap_html("", false, RemoteImages::AskToLoad);
        assert!(ask.contains("img-src https://mdview.example data:;"));
        assert!(ask.contains("var askRemoteImages = true"));
        let blocked = wrap_html("", false, RemoteImages::Blocked);
        assert!(blocked.contains("img-src https://mdview.example data:;"));
        assert!(blocked.contains("var askRemoteImages = false"));
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
