use pulldown_cmark::{Options, Parser, html};

pub fn markdown_to_html(markdown: &str) -> String {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);

    let parser = Parser::new_ext(markdown, options);

    let mut html_output = String::new();
    html::push_html(&mut html_output, parser);
    html_output
}

pub fn wrap_html(content: &str, dark_mode: bool) -> String {
    let bg_color = if dark_mode { "#1e1e1e" } else { "#ffffff" };
    let text_color = if dark_mode { "#d4d4d4" } else { "#24292e" };
    let code_bg = if dark_mode { "#2d2d2d" } else { "#f6f8fa" };
    let link_color = if dark_mode { "#58a6ff" } else { "#0366d6" };
    let border_color = if dark_mode { "#444" } else { "#e1e4e8" };

    format!(
        r#"<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
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
<script>
(function() {{
    function slugify(text) {{
        return (text || '')
            .toLowerCase()
            .trim()
            .replace(/[\s]+/g, '-')
            .replace(/[^a-z0-9\-_]/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }}

    function assignHeadingIds() {{
        var seen = Object.create(null);
        document.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach(function(heading) {{
            if (heading.id) return;

            var baseId = slugify(heading.textContent);
            if (!baseId) return;

            var id = baseId;
            var counter = 1;
            while (seen[id] || document.getElementById(id)) {{
                counter += 1;
                id = baseId + '-' + counter;
            }}

            seen[id] = true;
            heading.id = id;
        }});
    }}

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

    assignHeadingIds();
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

    let options = Options::empty();
    let parser = Parser::new_ext(markdown, options);

    let mut output = String::new();

    for event in parser {
        match event {
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

    #[test]
    fn test_basic_markdown() {
        let md = "# Hello\n\nThis is **bold** and *italic*.";
        let html = markdown_to_html(md);
        assert!(html.contains("<h1>Hello</h1>"));
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
}
