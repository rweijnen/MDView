use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd, html};
use std::collections::HashMap;

/// HTML for the viewer's page showing `document`: absolute paths in images
/// and links (`D:\x.png`, `file:///D:/x.png`, `\\server\share\x.png`) become
/// MDView URLs, so the host decides whether they may be loaded (see `vpath`).
pub fn markdown_to_page_html(markdown: &str, document: &std::path::Path) -> String {
    render(markdown, Some(document))
}

/// HTML without any rewriting (console output: `--html`, `--body`).
#[allow(dead_code)] // standalone viewer only; markdown.rs is compiled into both binaries
pub fn markdown_to_html(markdown: &str) -> String {
    render(markdown, None)
}

/// Absolute targets of images and links, as MDView URLs. Raw HTML is left
/// alone; its absolute paths stay blocked by the page's rules.
fn rewrite_absolute_targets(events: &mut [Event], document: &std::path::Path) {
    let Some(root) = crate::vpath::Root::of(document) else { return };
    for event in events.iter_mut() {
        if let Event::Start(Tag::Image { dest_url, .. } | Tag::Link { dest_url, .. }) = event {
            if let Some(url) = crate::vpath::absolute_link_target(dest_url)
                .and_then(|path| crate::vpath::url_for(&path, &root))
            {
                *dest_url = url.into();
            }
        }
    }
}

fn render(markdown: &str, document: Option<&std::path::Path>) -> String {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);
    options.insert(Options::ENABLE_YAML_STYLE_METADATA_BLOCKS);
    options.insert(Options::ENABLE_PLUSES_DELIMITED_METADATA_BLOCKS);

    let mut events: Vec<Event> = Parser::new_ext(markdown, options).collect();
    assign_heading_ids(&mut events);
    if let Some(document) = document {
        rewrite_absolute_targets(&mut events, document);
    }

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
    /// Only once the host blocked a file from a network location.
    WhenNetworkFileBlocked,
}

pub const REMOTE_IMAGES_BAR: &str = "remote-images";

/// Files from another share or a network drive: asked separately from remote
/// images, because a share can also receive the user's Windows login.
pub const NETWORK_FILES_BAR: &str = "network-files";

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
    /// Files from network locations are blocked until the user agrees.
    pub ask_network_files: bool,
    /// Bars the user closed while this document is open.
    pub closed: &'a [String],
}

/// Zoom levels in percent for Ctrl+mouse wheel, Ctrl+Plus and Ctrl+Minus
/// (#16): the same steps as Edge.
pub const ZOOM_STEPS: [u32; 17] = [25, 33, 50, 67, 75, 80, 90, 100, 110, 125, 150, 175, 200, 250, 300, 400, 500];
pub const DEFAULT_ZOOM: u32 = 100;

pub fn is_zoom_step(percent: u32) -> bool {
    ZOOM_STEPS.contains(&percent)
}

/// The zoom after the user's zoom key or wheel: `"in"`, `"out"` or `"reset"`.
/// `None` for anything else. At the ends it stays where it is.
pub fn next_zoom(current: u32, direction: &str) -> Option<u32> {
    match direction {
        "in" => Some(ZOOM_STEPS.iter().copied().find(|&s| s > current).unwrap_or(ZOOM_STEPS[ZOOM_STEPS.len() - 1])),
        "out" => Some(ZOOM_STEPS.iter().rev().copied().find(|&s| s < current).unwrap_or(ZOOM_STEPS[0])),
        "reset" => Some(DEFAULT_ZOOM),
        _ => None,
    }
}

/// JSON for `PostWebMessageAsJson`: the page applies this zoom.
pub fn zoom_message(percent: u32) -> String {
    format!(r#"{{"type":"zoom","percent":{percent}}}"#)
}

/// How MDView's pages look, from the user's settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageLook {
    /// Zoom of the whole page, one of `ZOOM_STEPS` (#16).
    pub zoom: u32,
    /// Fonts for the document (#50).
    pub fonts: crate::settings::FontSettings,
}

impl Default for PageLook {
    fn default() -> Self {
        PageLook { zoom: DEFAULT_ZOOM, fonts: crate::settings::FontSettings::default() }
    }
}

/// The settings panel (#50), part of the page script (it runs inside the
/// page's IIFE, where `webview` is set). MDView sends `showSettings` with the
/// saved settings and the installed fonts; the panel previews font changes
/// live through the page's CSS variables, and Save sends `saveSettings`.
/// MDView answers `settingsSaved` or `settingsError` (shown in the panel).
/// Built in a closed shadow root, so the document's CSS can't restyle or
/// relabel it; every text from MDView is set with `textContent`.
const SETTINGS_SCRIPT: &str = r##"
if (webview) (function() {
    var DEFAULT_TEXT = '-apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif';
    var DEFAULT_CODE = '"Cascadia Code", "Fira Code", Consolas, monospace';
    var VARS = ['--mdview-text-font', '--mdview-text-size', '--mdview-heading-font', '--mdview-code-font', '--mdview-code-size'];
    var PRESETS = [
        {px: 12, glyph: 'a', name: 'Small text', css: 'font-size:13px'},
        {px: 14, glyph: 'A', name: 'Normal text', css: 'font-size:16px'},
        {px: 17, glyph: 'A', name: 'Large text', css: 'font-size:20px;font-weight:700'}
    ];
    var TEXT_SIZES = [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 22, 24, 28, 32, 40, 48];
    var CODE_SIZES = [0, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24];
    var root = document.documentElement.style;
    var open = null;

    var CSS = [
        ':host{all:initial}',
        '.scrim{position:fixed;inset:0;z-index:2147483647;display:flex;align-items:flex-start;justify-content:center;padding:32px 16px;box-sizing:border-box;background:rgba(27,31,36,.35);overflow:auto;font-family:"Segoe UI",-apple-system,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.45;color:#24292e}',
        '.dark.scrim{background:rgba(0,0,0,.55);color:#d4d4d4}',
        '.box{width:560px;max-width:100%;background:#fff;border:1px solid #d0d7de;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,.28)}',
        '.dark .box{background:#252526;border-color:#454545}',
        '.head{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid #d0d7de}',
        '.dark .head,.dark .foot{border-color:#454545}',
        '.head h2{margin:0;font-size:18px;font-weight:600}',
        '.muted{font-size:12px;color:#57606a}.dark .muted{color:#a3abb4}',
        '.body{display:flex;flex-direction:column;gap:22px;padding:20px}',
        'section{display:flex;flex-direction:column;gap:8px}',
        'h3{margin:0;font-size:14px;font-weight:600}',
        '.row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}',
        'button,select,input{font:inherit;color:inherit}',
        'button{cursor:pointer}',
        '.seg{min-height:36px;min-width:44px;padding:4px 14px;border-radius:4px;border:1px solid #d0d7de;background:#fff}',
        '.dark .seg{border-color:#454545;background:#252526}',
        '.seg[aria-pressed="true"]{border:2px solid #0969da;background:#ddf4ff;color:#0550ae;font-weight:600}',
        '.dark .seg[aria-pressed="true"]{border-color:#4493f8;background:#13294a;color:#cfe3ff}',
        '.btn{min-height:32px;min-width:88px;padding:4px 14px;border-radius:4px;border:1px solid #d0d7de;background:#fff}',
        '.dark .btn{border-color:#454545;background:#252526}',
        '.primary{background:#0969da;border-color:#0969da;color:#fff;font-weight:600}',
        '.dark .primary{background:#1f6feb;border-color:#1f6feb;color:#fff}',
        '.primary:disabled{opacity:.6;cursor:default}',
        '.close{width:32px;height:32px;border:none;border-radius:4px;background:transparent;display:flex;align-items:center;justify-content:center}',
        '.adv{align-self:flex-start;display:flex;align-items:center;gap:6px;min-height:32px;padding:4px 0;border:none;background:transparent;font-weight:600}',
        '.adv svg{transition:transform .15s}.adv[aria-expanded="true"] svg{transform:rotate(90deg)}',
        '.grid{display:grid;grid-template-columns:96px minmax(0,1fr) 150px;gap:10px 12px;align-items:center;padding:12px;border:1px solid #d0d7de;border-radius:6px;background:#f6f8fa}',
        '.dark .grid{border-color:#454545;background:#1e1e1e}',
        'select{min-height:32px;width:100%;box-sizing:border-box;padding:4px 8px;border:1px solid #d0d7de;border-radius:4px;background:#fff;font-size:14px}',
        '.dark select{border-color:#454545;background:#252526}',
        '.sep{height:1px;background:#d0d7de}.dark .sep{background:#454545}',
        '.check{display:flex;align-items:flex-start;gap:10px;cursor:pointer;min-height:32px}',
        '.check input{width:18px;height:18px;margin:1px 0 0;accent-color:#0969da}',
        '.dark .check input{accent-color:#4493f8}',
        '.stack{display:flex;flex-direction:column;gap:2px}',
        '.strong{font-weight:600}',
        '.notes{padding:10px 12px;border:1px solid #d4a72c;border-radius:6px;background:#fff8c5;font-size:13px}',
        '.dark .notes{border-color:#6b5b2a;background:#3b3420}',
        '.error{margin:0 20px 4px;padding:10px 12px;border:1px solid #ff8182;border-radius:6px;background:#ffebe9;color:#82071e;font-size:13px;white-space:pre-line;overflow-wrap:anywhere}',
        '.dark .error{border-color:#a4282f;background:#3a1d1f;color:#ffc1bd}',
        '.foot{display:flex;justify-content:flex-end;gap:8px;padding:14px 20px;border-top:1px solid #d0d7de}',
        ':focus-visible{outline:2px solid #0969da;outline-offset:2px}.dark :focus-visible{outline-color:#4493f8}',
        '[hidden]{display:none!important}'
    ].join('\n');

    function el(tag, props, children) {
        var e = document.createElement(tag);
        Object.keys(props || {}).forEach(function(k) {
            var v = props[k];
            if (k === 'text') e.textContent = v;
            else if (k === 'className') e.className = v;
            else if (k === 'style') e.style.cssText = v;
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), v);
            else if (v !== null && v !== undefined && v !== false) e.setAttribute(k, v === true ? '' : String(v));
        });
        (children || []).forEach(function(c) { if (c) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
        return e;
    }
    function svg(path, size) {
        var ns = 'http://www.w3.org/2000/svg';
        var s = document.createElementNS(ns, 'svg');
        s.setAttribute('width', size); s.setAttribute('height', size);
        s.setAttribute('viewBox', '0 0 14 14'); s.setAttribute('fill', 'none');
        s.setAttribute('stroke', 'currentColor'); s.setAttribute('stroke-width', '1.6');
        s.setAttribute('stroke-linecap', 'round'); s.setAttribute('stroke-linejoin', 'round');
        s.setAttribute('aria-hidden', 'true');
        var p = document.createElementNS(ns, 'path'); p.setAttribute('d', path); s.appendChild(p);
        return s;
    }
    function fontCss(name, stack) { return name ? '"' + name + '", ' + stack : stack; }

    function preview(s) {
        root.setProperty('--mdview-text-font', fontCss(s.textFont, DEFAULT_TEXT));
        root.setProperty('--mdview-text-size', s.textSize + 'px');
        root.setProperty('--mdview-heading-font', s.headingFont ? fontCss(s.headingFont, DEFAULT_TEXT) : 'var(--mdview-text-font)');
        root.setProperty('--mdview-code-font', fontCss(s.codeFont, DEFAULT_CODE));
        root.setProperty('--mdview-code-size', s.codeSize ? s.codeSize + 'px' : '85%');
    }
    function unpreview() { VARS.forEach(function(v) { root.removeProperty(v); }); }

    function close(keepPreview) {
        if (!open) return;
        if (!keepPreview) unpreview();
        open.host.remove();
        panelClosed();
        var back = open.back;
        open = null;
        if (back && back.focus) back.focus();
    }

    function show(m) {
        if (open || panels) return;
        var s = {};
        Object.keys(m.settings).forEach(function(k) { s[k] = m.settings[k]; });
        var fonts = (m.fonts || []).map(function(f) { return f.name; });
        var monoFonts = (m.fonts || []).filter(function(f) { return f.mono; }).map(function(f) { return f.name; });
        var host = el('div', {'data-mdview-settings': ''});
        var shadow = host.attachShadow({mode: 'closed'});
        shadow.appendChild(el('style', {text: CSS}));

        function seg(label, pressed, onclick, extra) {
            return el('button', {type: 'button', className: 'seg', 'aria-pressed': pressed ? 'true' : 'false', 'aria-label': extra && extra.name, style: extra && extra.css, onclick: onclick}, [label]);
        }
        function select(label, value, options, onchange, font) {
            var sel = el('select', {'aria-label': label, onchange: function() { onchange(sel.value); }, style: font ? 'font-family:' + font : null});
            var found = false;
            options.forEach(function(o) {
                var opt = el('option', {value: o.value, style: o.font ? 'font-family:' + o.font : null}, [o.label]);
                if (o.value === String(value)) { opt.selected = true; found = true; }
                sel.appendChild(opt);
            });
            if (!found) {
                var keep = el('option', {value: String(value)}, [String(value) + ' (not installed)']);
                keep.selected = true;
                sel.appendChild(keep);
            }
            return sel;
        }
        function fontOptions(list, defaultLabel) {
            return [{value: '', label: defaultLabel}].concat(list.map(function(n) {
                return {value: n, label: n, font: fontCss(n, 'sans-serif')};
            }));
        }

        var sizeRow = el('div', {className: 'row', role: 'group', 'aria-label': 'Font size'});
        var custom = el('span', {className: 'muted', text: 'Custom (set under Advanced)'});
        var advBody = el('div', {});
        var advButton = el('button', {type: 'button', className: 'adv', 'aria-expanded': 'false'}, [svg('M5 3l4 4-4 4', 12), 'Advanced: fonts per kind of text']);
        var askRows = el('div', {style: 'display:flex;flex-direction:column;gap:22px'});
        var error = el('div', {className: 'error', role: 'alert', hidden: true});
        var save = el('button', {type: 'button', className: 'btn primary', text: 'Save'});
        var cancel = el('button', {type: 'button', className: 'btn', text: 'Cancel'});

        function renderSizes() {
            sizeRow.textContent = '';
            PRESETS.forEach(function(p) {
                sizeRow.appendChild(seg(p.glyph, s.textSize === p.px, function() { s.textSize = p.px; changed(); }, p));
            });
            custom.hidden = PRESETS.some(function(p) { return p.px === s.textSize; });
            sizeRow.appendChild(custom);
        }
        function renderAdvanced() {
            advBody.textContent = '';
            if (advButton.getAttribute('aria-expanded') !== 'true') return;
            var grid = el('div', {className: 'grid'}, [
                el('span', {}), el('span', {className: 'muted', text: 'Font'}), el('span', {className: 'muted', text: 'Size'}),
                el('span', {className: 'strong', text: 'Text'}),
                select('Text font', s.textFont, fontOptions(fonts, 'Segoe UI (default)'), function(v) { s.textFont = v; changed(); }, fontCss(s.textFont, DEFAULT_TEXT)),
                select('Text size', s.textSize, TEXT_SIZES.map(function(n) { return {value: String(n), label: n + ' px' + (n === 14 ? ' (default)' : '')}; }), function(v) { s.textSize = Number(v); changed(); }),
                el('span', {className: 'strong', text: 'Headings'}),
                select('Heading font', s.headingFont, fontOptions(fonts, 'Same as Text'), function(v) { s.headingFont = v; changed(); }, s.headingFont ? fontCss(s.headingFont, DEFAULT_TEXT) : null),
                el('span', {className: 'muted', text: 'in step with Text'}),
                el('span', {className: 'strong', text: 'Code'}),
                select('Code font', s.codeFont, fontOptions(monoFonts.indexOf(s.codeFont) < 0 && s.codeFont ? monoFonts.concat([s.codeFont]) : monoFonts, 'Cascadia Code (default)'), function(v) { s.codeFont = v; changed(); }, fontCss(s.codeFont, DEFAULT_CODE)),
                select('Code size', s.codeSize, CODE_SIZES.map(function(n) { return {value: String(n), label: n === 0 ? '85% of Text (default)' : n + ' px'}; }), function(v) { s.codeSize = Number(v); changed(); }),
                el('span', {}),
                el('span', {className: 'muted', style: 'grid-column:span 2', text: 'Fonts installed on this PC. Changes show in the document behind this panel.'})
            ]);
            advBody.appendChild(el('div', {style: 'display:flex;flex-direction:column;gap:10px'}, [
                grid,
                el('button', {type: 'button', className: 'btn', style: 'align-self:flex-start', text: 'Reset fonts to default', onclick: function() {
                    s.textFont = ''; s.textSize = 14; s.headingFont = ''; s.codeFont = ''; s.codeSize = 0; changed();
                }})
            ]));
        }
        function renderAsk() {
            askRows.textContent = '';
            [['remoteImages', 'Remote images', 'Images from the internet in a document.'],
             ['networkFiles', 'Network files', 'Images and files from another share or a network drive.']].forEach(function(r) {
                var row = el('div', {className: 'row', role: 'group', 'aria-label': r[1]});
                [['ask', 'Ask'], ['always', 'Always load'], ['never', 'Never']].forEach(function(o) {
                    row.appendChild(seg(o[1], s[r[0]] === o[0], function() { s[r[0]] = o[0]; renderAsk(); }));
                });
                askRows.appendChild(el('section', {}, [el('div', {className: 'stack'}, [el('h3', {text: r[1]}), el('span', {className: 'muted', text: r[2]})]), row]));
            });
        }
        function check(key, title, help) {
            var input = el('input', {type: 'checkbox', onchange: function() { s[key] = input.checked; }});
            input.checked = !!s[key];
            return el('label', {className: 'check'}, [input, el('span', {className: 'stack'}, [el('span', {className: 'strong', text: title}), el('span', {className: 'muted', text: help})])]);
        }
        function changed() {
            preview(s);
            renderSizes();
            var focused = shadow.activeElement && shadow.activeElement.getAttribute('aria-label');
            renderAdvanced();
            if (focused) {
                var again = advBody.querySelector('[aria-label="' + focused + '"]');
                if (again) again.focus();
            }
        }
        advButton.addEventListener('click', function() {
            advButton.setAttribute('aria-expanded', advButton.getAttribute('aria-expanded') === 'true' ? 'false' : 'true');
            renderAdvanced();
        });

        var notes = (m.notes || []).length ? el('div', {className: 'notes'}, (m.notes || []).map(function(n) { return el('div', {text: n}); })) : null;
        var box = el('div', {className: 'box', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'mdview-settings-title'}, [
            el('div', {className: 'head'}, [
                el('div', {className: 'stack', style: 'flex:1'}, [
                    el('h2', {id: 'mdview-settings-title', text: 'Settings'}),
                    el('span', {className: 'muted', text: m.program})
                ]),
                el('button', {type: 'button', className: 'close', 'aria-label': 'Close', title: 'Close (Esc)', onclick: function() { close(false); }}, [svg('M3 3l8 8M11 3l-8 8', 14)])
            ]),
            el('div', {className: 'body'}, [
                notes,
                el('section', {}, [
                    el('div', {className: 'row', style: 'align-items:baseline'}, [el('h3', {text: 'Font size'}), el('span', {className: 'muted', text: 'all text in the document'})]),
                    sizeRow
                ]),
                el('section', {}, [advButton, advBody]),
                el('div', {className: 'sep'}),
                askRows,
                check('askDownloads', 'Ask before showing a downloaded document', 'Documents with the "Mark of the Web" from a browser or mail program.'),
                s.askRegister === null ? null : check('askRegister', 'Offer to register MDView as .md viewer', 'The bar at the top of a document. Standalone viewer only.')
            ]),
            error,
            el('div', {className: 'foot'}, [cancel, save])
        ]);
        var scrim = el('div', {className: 'scrim' + (m.dark ? ' dark' : '')}, [box]);
        shadow.appendChild(scrim);

        cancel.addEventListener('click', function() { close(false); });
        save.addEventListener('click', function() {
            error.hidden = true;
            save.disabled = true;
            save.textContent = 'Saving…';
            var msg = {
                type: 'saveSettings',
                textFont: s.textFont, textSize: String(s.textSize), headingFont: s.headingFont,
                codeFont: s.codeFont, codeSize: String(s.codeSize),
                remoteImages: s.remoteImages, networkFiles: s.networkFiles,
                askDownloads: s.askDownloads ? 'true' : 'false'
            };
            // Only the standalone viewer has the register setting.
            if (s.askRegister !== null) msg.askRegister = s.askRegister ? 'true' : 'false';
            webview.postMessage(msg);
        });

        // Keyboard: Esc cancels, Tab stays inside the panel.
        scrim.addEventListener('keydown', function(e) {
            if (e.key !== 'Tab') return;
            var items = Array.prototype.filter.call(box.querySelectorAll('button,select,input'), function(x) { return !x.disabled && x.offsetParent !== null; });
            if (!items.length) return;
            var first = items[0], last = items[items.length - 1], at = shadow.activeElement;
            if (e.shiftKey && at === first) { e.preventDefault(); last.focus(); }
            else if (!e.shiftKey && at === last) { e.preventDefault(); first.focus(); }
        });

        renderSizes();
        renderAdvanced();
        renderAsk();
        open = {
            host: host, back: document.activeElement,
            saved: function() { close(true); },
            failed: function(text) {
                error.textContent = text;
                error.hidden = false;
                save.disabled = false;
                save.textContent = 'Save';
            },
            values: s
        };
        document.body.appendChild(host);
        panelOpened();
        var start = sizeRow.querySelector('[aria-pressed="true"]') || sizeRow.querySelector('button');
        if (start) start.focus();
    }

    window.addEventListener('keydown', function(e) {
        if (open && e.key === 'Escape') {
            e.preventDefault();
            e.stopImmediatePropagation();
            close(false);
        } else if (!panels && e.ctrlKey && !e.altKey && e.key === ',') {
            e.preventDefault();
            webview.postMessage({type: 'openSettings'});
        }
    }, true);
    webview.addEventListener('message', function(e) {
        var m = e.data;
        if (!m || typeof m.type !== 'string') return;
        if (m.type === 'showSettings' && m.settings) show(m);
        else if (m.type === 'settingsSaved' && open) open.saved();
        else if (m.type === 'settingsError' && open && typeof m.message === 'string') open.failed(m.message);
    });
})();
"##;

/// Shared by the panels (settings, help; #50, #59): only one is open at a
/// time, and MDView is told whether one is open (no right-click menu then).
/// Also passes N / P / 1-9 to Total Commander's Lister when the focus is in
/// the document itself (not in a panel or a field); the standalone viewer
/// ignores them.
const PANEL_SCRIPT: &str = r##"
    var panels = 0;
    function panelOpened() {
        panels++;
        if (webview) webview.postMessage({type: 'panel', open: 'true'});
    }
    function panelClosed() {
        panels = Math.max(0, panels - 1);
        if (webview) webview.postMessage({type: 'panel', open: panels ? 'true' : 'false'});
    }
    if (webview) document.addEventListener('keydown', function(e) {
        if (panels || e.ctrlKey || e.altKey || e.metaKey || e.key.length !== 1) return;
        if (!/^[NnPp1-9]$/.test(e.key)) return;
        var t = e.target;
        if (t && (t.isContentEditable || /^(INPUT|SELECT|TEXTAREA)$/.test(t.tagName))) return;
        e.preventDefault();
        webview.postMessage({type: 'listerKey', key: e.key.toUpperCase()});
    });
"##;

/// The keyboard help (#59): F1 asks MDView, which sends `showHelp` with this
/// program's keys. A panel like the settings panel (closed shadow root, text
/// with `textContent`); Esc, F1 or Close closes it.
const HELP_SCRIPT: &str = r##"
if (webview) (function() {
    var open = null;
    var CSS = [
        ':host{all:initial}',
        '.scrim{position:fixed;inset:0;z-index:2147483647;display:flex;align-items:flex-start;justify-content:center;padding:32px 16px;box-sizing:border-box;background:rgba(27,31,36,.35);overflow:auto;font-family:"Segoe UI",-apple-system,Helvetica,Arial,sans-serif;font-size:14px;line-height:1.45;color:#24292e}',
        '.dark.scrim{background:rgba(0,0,0,.55);color:#d4d4d4}',
        '.box{width:560px;max-width:100%;background:#fff;border:1px solid #d0d7de;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,.28)}',
        '.dark .box{background:#252526;border-color:#454545}',
        '.head{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid #d0d7de}',
        '.dark .head{border-color:#454545}',
        '.head h2{margin:0;font-size:18px;font-weight:600}',
        '.muted{font-size:12px;color:#57606a}.dark .muted{color:#a3abb4}',
        '.stack{display:flex;flex-direction:column;gap:2px}',
        '.body{display:flex;flex-direction:column;gap:18px;padding:16px 20px 20px}',
        'h3{margin:0 0 6px;font-size:14px;font-weight:600}',
        'table{border-collapse:collapse;width:100%}',
        'td{padding:5px 0;border-top:1px solid #eaeef2;vertical-align:top}',
        '.dark td{border-color:#3a3a3a}',
        'td.key{width:46%;padding-right:12px}',
        'kbd{display:inline-block;padding:1px 6px;border:1px solid #d0d7de;border-bottom-width:2px;border-radius:4px;background:#f6f8fa;font-family:"Segoe UI",sans-serif;font-size:12px}',
        '.dark kbd{border-color:#555;background:#1e1e1e}',
        'button{font:inherit;color:inherit;cursor:pointer}',
        '.close{width:32px;height:32px;border:none;border-radius:4px;background:transparent;display:flex;align-items:center;justify-content:center}',
        ':focus-visible{outline:2px solid #0969da;outline-offset:2px}.dark :focus-visible{outline-color:#4493f8}'
    ].join('\n');

    function el(tag, props, children) {
        var e = document.createElement(tag);
        Object.keys(props || {}).forEach(function(k) {
            if (k === 'text') e.textContent = props[k];
            else if (k === 'className') e.className = props[k];
            else if (k.indexOf('on') === 0) e.addEventListener(k.slice(2), props[k]);
            else e.setAttribute(k, props[k]);
        });
        (children || []).forEach(function(c) { if (c) e.appendChild(c); });
        return e;
    }

    function close() {
        if (!open) return;
        open.host.remove();
        var back = open.back;
        open = null;
        panelClosed();
        if (back && back.focus) back.focus();
    }

    function show(m) {
        if (open || panels) return;
        var host = el('div', {'data-mdview-help': ''});
        var shadow = host.attachShadow({mode: 'closed'});
        shadow.appendChild(el('style', {text: CSS}));
        var ns = 'http://www.w3.org/2000/svg';
        var x = document.createElementNS(ns, 'svg');
        x.setAttribute('width', '14'); x.setAttribute('height', '14'); x.setAttribute('viewBox', '0 0 14 14');
        x.setAttribute('fill', 'none'); x.setAttribute('stroke', 'currentColor'); x.setAttribute('stroke-width', '1.6');
        x.setAttribute('stroke-linecap', 'round'); x.setAttribute('aria-hidden', 'true');
        var p = document.createElementNS(ns, 'path'); p.setAttribute('d', 'M3 3l8 8M11 3l-8 8'); x.appendChild(p);
        var closeButton = el('button', {type: 'button', className: 'close', 'aria-label': 'Close', title: 'Close (Esc)', onclick: close}, [x]);
        var body = el('div', {className: 'body'});
        (m.groups || []).forEach(function(g) {
            var table = el('table', {});
            (g.keys || []).forEach(function(k) {
                table.appendChild(el('tr', {}, [
                    el('td', {className: 'key'}, [el('kbd', {text: String(k[0])})]),
                    el('td', {text: String(k[1])})
                ]));
            });
            body.appendChild(el('section', {}, [el('h3', {text: String(g.title)}), table]));
        });
        var box = el('div', {className: 'box', role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'mdview-help-title'}, [
            el('div', {className: 'head'}, [
                el('div', {className: 'stack', style: 'flex:1'}, [
                    el('h2', {id: 'mdview-help-title', text: 'Keyboard shortcuts'}),
                    el('span', {className: 'muted', text: String(m.program)})
                ]),
                closeButton
            ]),
            body
        ]);
        var scrim = el('div', {className: 'scrim' + (m.dark ? ' dark' : '')}, [box]);
        scrim.addEventListener('mousedown', function(e) { if (e.target === scrim) close(); });
        shadow.appendChild(scrim);
        open = {host: host, back: document.activeElement};
        document.body.appendChild(host);
        panelOpened();
        closeButton.focus();
    }

    window.addEventListener('keydown', function(e) {
        if (e.key !== 'F1' && !(open && e.key === 'Escape')) return;
        e.preventDefault();
        e.stopImmediatePropagation();
        if (open) close();
        else if (!panels) webview.postMessage({type: 'openHelp'});
    }, true);
    webview.addEventListener('message', function(e) {
        var m = e.data;
        if (m && m.type === 'showHelp' && Array.isArray(m.groups)) show(m);
    });
})();
"##;

/// MDView's own fonts, used when no font is chosen and after a chosen one.
pub const DEFAULT_TEXT_STACK: &str = r#"-apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif"#;
pub const DEFAULT_CODE_STACK: &str = r#""Cascadia Code", "Fira Code", Consolas, monospace"#;

/// A CSS font-family value: the chosen font first, then MDView's defaults.
/// Only a valid font name (`settings::is_font_name`) is ever put in the CSS.
fn font_css(name: &str, stack: &str) -> String {
    if crate::settings::is_font_name(name) { format!("\"{name}\", {stack}") } else { stack.to_string() }
}

/// The page's font variables (#50), for its `:root` rule.
fn font_variables(fonts: &crate::settings::FontSettings) -> String {
    let text_size = if crate::settings::is_font_size(fonts.text_size) { fonts.text_size } else { crate::settings::DEFAULT_TEXT_SIZE };
    let heading = if fonts.heading_font.is_empty() { "var(--mdview-text-font)".to_string() } else { font_css(&fonts.heading_font, DEFAULT_TEXT_STACK) };
    let code_size = if crate::settings::is_font_size(fonts.code_size) { format!("{}px", fonts.code_size) } else { "85%".to_string() };
    format!(
        "--mdview-text-font: {}; --mdview-text-size: {text_size}px; --mdview-heading-font: {heading}; --mdview-code-font: {}; --mdview-code-size: {code_size};",
        font_css(&fonts.text_font, DEFAULT_TEXT_STACK),
        font_css(&fonts.code_font, DEFAULT_CODE_STACK),
    )
}

/// CSS `zoom` value for a percentage: 125 gives `1.25`.
fn css_zoom(percent: u32) -> String {
    let whole = percent / 100;
    let fraction = percent % 100;
    if fraction == 0 {
        whole.to_string()
    } else {
        format!("{whole}.{}", format!("{fraction:02}").trim_end_matches('0'))
    }
}

/// What the user chose while a document is open (#24): bars they closed, and
/// whether they opened it although it was downloaded from the internet.
/// Reloading the same document keeps these; another document starts fresh.
#[derive(Debug, Default)]
pub struct DocumentChoices {
    document: Option<std::path::PathBuf>,
    closed_bars: Vec<String>,
    opened_download: bool,
    network_loaded: bool,
    /// Network locations blocked on the page now shown, in the order met.
    network_blocked: Vec<String>,
}

impl DocumentChoices {
    /// Switch to `document`; choices made for another document are dropped.
    fn switch_to(&mut self, document: &std::path::Path) {
        if self.document.as_deref() != Some(document) {
            self.document = Some(document.to_path_buf());
            self.closed_bars.clear();
            self.opened_download = false;
            self.network_loaded = false;
            self.network_blocked.clear();
        }
    }

    /// The document whose page is shown now (the last one served).
    pub fn document(&self) -> Option<&std::path::Path> {
        self.document.as_deref()
    }

    /// `document` is served as a page (again): what was blocked on the
    /// previous page no longer counts.
    pub fn page_served(&mut self, document: &std::path::Path) {
        self.switch_to(document);
        self.network_blocked.clear();
    }

    /// Whether the user chose Load files for `document` while it is open.
    pub fn network_loaded(&mut self, document: &std::path::Path) -> bool {
        self.switch_to(document);
        self.network_loaded
    }

    /// The user chose Load files for `document`.
    pub fn load_network(&mut self, document: &std::path::Path) {
        self.switch_to(document);
        self.network_loaded = true;
        self.network_blocked.clear();
    }

    /// A file from `location` was blocked on `document`'s page. True the
    /// first time for that location, so the page is told only once.
    pub fn note_network_blocked(&mut self, document: &std::path::Path, location: &str) -> bool {
        self.switch_to(document);
        if self.network_blocked.iter().any(|l| l.eq_ignore_ascii_case(location)) {
            return false;
        }
        self.network_blocked.push(location.to_string());
        true
    }

    /// The network locations blocked on `document`'s page so far.
    pub fn network_blocked(&mut self, document: &std::path::Path) -> Vec<String> {
        self.switch_to(document);
        self.network_blocked.clone()
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
    if page.ask_network_files {
        bars.push(Bar {
            id: NETWORK_FILES_BAR,
            text: "This document loads files from a network location:".to_string(),
            buttons: vec![
                BarButton { choice: "once", label: "Load files", closes_bar: false },
                BarButton { choice: "always", label: "Always load", closes_bar: false },
                BarButton { choice: "never", label: "Never", closes_bar: true },
            ],
            show: BarShow::WhenNetworkFileBlocked,
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
pub fn script_string(s: &str) -> String {
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

/// The host's message to the page that a network location was blocked, as
/// JSON for `PostWebMessageAsJson`. The location is shown as text only.
pub fn network_blocked_message(location: &str) -> String {
    format!(r#"{{"type":"networkBlocked","location":{}}}"#, script_string(location))
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
                BarShow::WhenNetworkFileBlocked => "networkFileBlocked",
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
        RemoteImages::Allowed => "https://mdview.example https://*.mdview.example data: https: http:",
        RemoteImages::AskToLoad | RemoteImages::Blocked => "https://mdview.example https://*.mdview.example data:",
    };
    format!(
        "default-src 'none'; script-src 'nonce-{nonce}'; style-src 'unsafe-inline'; \
         img-src {img_src}; media-src https://mdview.example https://*.mdview.example; font-src 'none'; \
         connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; \
         form-action 'none'"
    )
}

pub fn wrap_html(content: &str, dark_mode: bool, remote_images: RemoteImages, bars: &[Bar], look: &PageLook) -> String {
    let nonce = new_nonce();
    let zoom = css_zoom(look.zoom);
    let font_vars = font_variables(&look.fonts);
    let settings_script = SETTINGS_SCRIPT;
    let help_script = HELP_SCRIPT;
    let panel_script = PANEL_SCRIPT;
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
:root {{ zoom: {zoom}; {font_vars} }}
h1, h2, h3, h4, h5, h6 {{ font-family: var(--mdview-heading-font); }}
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
.mdview-gate code, .mdview-bar-location {{ word-break: break-all; }}
.mdview-bar button.mdview-bar-close {{
    border: none;
    background: transparent;
    color: inherit;
    font-size: 16px;
    line-height: 1;
    padding: 0 4px;
}}
body {{
    font-family: var(--mdview-text-font);
    font-size: var(--mdview-text-size);
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
    font-family: var(--mdview-code-font);
    font-size: var(--mdview-code-size);
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

    // MDView tells the page which network locations it blocked; they are
    // listed in the bar as text. Locations blocked before this script ran
    // are asked for once (networkStatus).
    var networkLocations = {{}};
    function networkFileBlocked(location) {{
        bars.forEach(function(def) {{
            if (def.show !== 'networkFileBlocked') return;
            showBar(def);
            var bar = document.querySelector('#mdview-bars > [data-bar="' + def.id + '"]');
            if (!bar || networkLocations[location.toLowerCase()]) return;
            networkLocations[location.toLowerCase()] = true;
            var item = document.createElement('code');
            item.className = 'mdview-bar-location';
            item.textContent = location;
            bar.firstChild.appendChild(document.createTextNode(' '));
            bar.firstChild.appendChild(item);
        }});
    }}
    if (webview && bars.some(function(def) {{ return def.show === 'networkFileBlocked'; }})) {{
        webview.addEventListener('message', function(e) {{
            var m = e.data;
            if (m && m.type === 'networkBlocked' && typeof m.location === 'string') networkFileBlocked(m.location);
        }});
        webview.postMessage({{type: 'networkStatus'}});
    }}

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

    // Zoom (#16): MDView picks the next step, saves it and tells the page,
    // so the zoom is part of every page MDView shows and survives links and
    // reloading. WebView2's own zoom is switched off.
    if (webview) {{
        function zoom(direction) {{
            webview.postMessage({{type: 'zoom', direction: direction}});
        }}
        document.addEventListener('keydown', function(e) {{
            if (!e.ctrlKey || e.altKey || e.metaKey) return;
            var direction = (e.key === '+' || e.key === '=') ? 'in'
                : (e.key === '-' || e.key === '_') ? 'out'
                : e.key === '0' ? 'reset' : null;
            if (!direction) return;
            e.preventDefault();
            zoom(direction);
        }});
        // One step per wheel notch; small touchpad deltas add up to a step.
        var wheel = 0;
        document.addEventListener('wheel', function(e) {{
            if (!e.ctrlKey) return;
            e.preventDefault();
            wheel += e.deltaY;
            if (Math.abs(wheel) >= 50) {{
                zoom(wheel < 0 ? 'in' : 'out');
                wheel = 0;
            }}
        }}, {{passive: false}});
        webview.addEventListener('message', function(e) {{
            var m = e.data;
            if (m && m.type === 'zoom' && typeof m.percent === 'number') {{
                document.documentElement.style.zoom = String(m.percent / 100);
            }}
        }});
    }}

    // Where a link goes (#60): MDView shows it in a Windows tooltip, worked
    // out the same way as a click. The link's own title (from the document)
    // is set aside meanwhile, so it can't show a second, different hint.
    if (webview) {{
        var hovered = null, hoveredTitle = null;
        function hoverEnd() {{
            if (!hovered) return;
            if (hoveredTitle !== null) hovered.setAttribute('title', hoveredTitle);
            hovered = null;
            hoveredTitle = null;
            webview.postMessage({{type: 'hoverLink', url: ''}});
        }}
        document.addEventListener('mouseover', function(e) {{
            var link = e.target && e.target.closest ? e.target.closest('a[href]') : null;
            if (link === hovered) return;
            hoverEnd();
            if (!link) return;
            hovered = link;
            hoveredTitle = link.getAttribute('title');
            if (hoveredTitle !== null) link.removeAttribute('title');
            webview.postMessage({{type: 'hoverLink', url: link.href, href: link.getAttribute('href')}});
        }});
        document.addEventListener('mouseout', function(e) {{
            if (hovered && !(e.relatedTarget && hovered.contains(e.relatedTarget))) hoverEnd();
        }});
        document.addEventListener('mousedown', hoverEnd, true);
        window.addEventListener('scroll', hoverEnd, true);
        window.addEventListener('blur', hoverEnd);
    }}
{panel_script}
{settings_script}
{help_script}
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

    /// A page at the default look (zoom 100%).
    fn wrap(content: &str, dark_mode: bool, remote_images: RemoteImages, bars: &[Bar]) -> String {
        wrap_html(content, dark_mode, remote_images, bars, &PageLook::default())
    }

    // #16: the zoom steps MDView moves through, and the ends.
    #[test]
    fn test_zoom_steps_up_down_and_reset() {
        assert_eq!(next_zoom(100, "in"), Some(110));
        assert_eq!(next_zoom(100, "out"), Some(90));
        assert_eq!(next_zoom(250, "reset"), Some(100));
        assert_eq!(next_zoom(500, "in"), Some(500), "stays at the largest step");
        assert_eq!(next_zoom(25, "out"), Some(25), "stays at the smallest step");
        assert_eq!(next_zoom(105, "in"), Some(110), "a value between steps goes to the next step");
        assert_eq!(next_zoom(105, "out"), Some(100));
        assert_eq!(next_zoom(100, "sideways"), None);
        let mut z = 25;
        for _ in 0..ZOOM_STEPS.len() {
            assert!(is_zoom_step(z));
            z = next_zoom(z, "in").unwrap();
        }
        assert_eq!(z, 500);
        assert!(!is_zoom_step(0) && !is_zoom_step(105) && !is_zoom_step(1000));
    }

    // #16: the zoom is part of the page, so navigating or reloading keeps it.
    #[test]
    fn test_page_carries_the_zoom() {
        assert!(wrap("", false, RemoteImages::Blocked, &[]).contains(":root { zoom: 1; "));
        for (percent, css) in [(125, "1.25"), (110, "1.1"), (33, "0.33"), (50, "0.5"), (400, "4")] {
            let page = wrap_html("", false, RemoteImages::Blocked, &[], &PageLook { zoom: percent, ..PageLook::default() });
            assert!(page.contains(&format!(":root {{ zoom: {css}; ")), "{percent}");
        }
        assert_eq!(zoom_message(125), r#"{"type":"zoom","percent":125}"#);
    }

    // #50: the page's fonts come from the settings as CSS variables; only a
    // valid font name gets into the CSS, never anything that could end it.
    #[test]
    fn test_page_fonts_come_from_the_settings() {
        use crate::settings::FontSettings;
        let page = wrap("", false, RemoteImages::Blocked, &[]);
        assert!(page.contains(&format!("--mdview-text-font: {DEFAULT_TEXT_STACK};")), "default text font");
        assert!(page.contains("--mdview-text-size: 14px;"));
        assert!(page.contains("--mdview-heading-font: var(--mdview-text-font);"));
        assert!(page.contains(&format!("--mdview-code-font: {DEFAULT_CODE_STACK};")));
        assert!(page.contains("--mdview-code-size: 85%;"));
        assert!(page.contains("font-family: var(--mdview-text-font);") && page.contains("font-family: var(--mdview-code-font);"));
        let fonts = FontSettings {
            text_font: "Georgia".into(), text_size: 17, heading_font: "Cambria".into(),
            code_font: "Consolas".into(), code_size: 13,
        };
        let page = wrap_html("", false, RemoteImages::Blocked, &[], &PageLook { fonts, ..PageLook::default() });
        assert!(page.contains(&format!("--mdview-text-font: \"Georgia\", {DEFAULT_TEXT_STACK};")));
        assert!(page.contains("--mdview-text-size: 17px;"));
        assert!(page.contains(&format!("--mdview-heading-font: \"Cambria\", {DEFAULT_TEXT_STACK};")));
        assert!(page.contains(&format!("--mdview-code-font: \"Consolas\", {DEFAULT_CODE_STACK};")));
        assert!(page.contains("--mdview-code-size: 13px;"));
        let evil = FontSettings { text_font: "x\"; } body { display:none".into(), text_size: 400, ..FontSettings::default() };
        let page = wrap_html("", false, RemoteImages::Blocked, &[], &PageLook { fonts: evil, ..PageLook::default() });
        assert!(!page.contains("x\"; } body") && page.contains("--mdview-text-size: 14px;"));
    }

    fn nonce_of(page: &str) -> String {
        let start = page.find("'nonce-").expect("CSP has a nonce") + 7;
        page[start..start + 32].to_string()
    }

    // Security: a document's own <script> and on...= handlers must not run;
    // only MDView's scripts carry the CSP nonce.
    #[test]
    fn test_page_csp_only_lets_mdview_scripts_run() {
        let body = markdown_to_html("# T\n\n<script>evil()</script>\n\n<img src=x onerror=\"evil()\">");
        let page = wrap(&body, false, RemoteImages::AskToLoad, &document_bars(&page(RemoteImages::AskToLoad, false)));
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
        let a = nonce_of(&wrap("", false, RemoteImages::Blocked, &[]));
        let b = nonce_of(&wrap("", false, RemoteImages::Blocked, &[]));
        assert_ne!(a, b);
    }

    // Absolute paths in Markdown images and links become MDView URLs, so the
    // host decides about them; raw HTML and console output are not rewritten.
    #[test]
    fn test_network_bar_is_offered_only_when_asked() {
        assert!(document_bars(&page(RemoteImages::Blocked, false)).is_empty());
        let bars = document_bars(&PageBars { ask_network_files: true, ..page(RemoteImages::AskToLoad, false) });
        let ids: Vec<&str> = bars.iter().map(|b| b.id).collect();
        assert_eq!(ids, [REMOTE_IMAGES_BAR, NETWORK_FILES_BAR], "network bar comes after remote images");
        let network = &bars[1];
        assert_eq!(network.show, BarShow::WhenNetworkFileBlocked);
        assert!(network.closable);
        for choice in ["once", "always", "never", CLOSE_CHOICE] {
            assert!(bar_choice_offered(&bars, NETWORK_FILES_BAR, choice), "{choice}");
        }
        assert!(!bar_choice_offered(&bars, NETWORK_FILES_BAR, "load"));
        assert!(network.buttons.iter().any(|b| b.choice == "never" && b.closes_bar));
        let page_html = wrap("x", false, RemoteImages::AskToLoad, &bars);
        assert!(page_html.contains(r#""show":"networkFileBlocked""#));
        // A closed network bar stays away while the document is open.
        let closed = vec![NETWORK_FILES_BAR.to_string()];
        let bars = document_bars(&PageBars { ask_network_files: true, closed: &closed, ..page(RemoteImages::Blocked, false) });
        assert!(bars.is_empty());
    }

    #[test]
    fn test_network_blocked_message_is_json_with_the_location_as_text() {
        assert_eq!(
            network_blocked_message(r"\\server\share"),
            r#"{"type":"networkBlocked","location":"\\\\server\\share"}"#
        );
        let evil = network_blocked_message(r#"\\x\"</script><b>"#);
        assert!(!evil.contains('<') && !evil.contains('>'), "{evil}");
    }

    #[test]
    fn test_network_choices_last_while_the_document_is_open() {
        let a = std::path::Path::new(r"C:\docs\a.md");
        let b = std::path::Path::new(r"C:\docs\b.md");
        let mut choices = DocumentChoices::default();
        choices.page_served(a);
        assert!(!choices.network_loaded(a));
        assert!(choices.note_network_blocked(a, r"\\server\share"));
        assert!(!choices.note_network_blocked(a, r"\\SERVER\share"), "told once per location");
        assert!(choices.note_network_blocked(a, "Z:"));
        assert_eq!(choices.network_blocked(a), [r"\\server\share", "Z:"]);
        choices.page_served(a);
        assert!(choices.network_blocked(a).is_empty(), "a new page starts without blocked locations");
        choices.load_network(a);
        assert!(choices.network_loaded(a));
        choices.page_served(a);
        assert!(choices.network_loaded(a), "a reload keeps Load files");
        choices.page_served(b);
        assert!(!choices.network_loaded(b), "another document asks again");
        assert!(!choices.network_loaded(a), "and so does the first one when it is opened again");
    }
    #[test]
    fn test_absolute_paths_become_mdview_urls() {
        let doc = std::path::Path::new(r"C:\docs\note.md");
        let md = concat!(
            "![a](D:\\pics\\a.png) ![b](<file:///D:/my%20pics/b.png>) ![c](\\\\\\\\server\\share\\c.png) ![c2](file://server/share/c2.png)\n\n",
            // In Markdown `\\` is an escaped backslash: this is `\server\share`, not a share.
            "![x](\\\\server\\share\\x.png)\n\n",
            "![d](C:\\docs\\img\\d.png) ![e](img/e.png) [f](D:\\notes\\f.md) [g](https://example.com/g)\n\n",
            "![p](\\\\.\\pipe\\p)\n\n",
            "<img src=\"D:\\raw.png\">\n"
        );
        let html = markdown_to_page_html(md, doc);
        for expected in [
            r#"src="https://drive-d.mdview.example/pics/a.png""#,
            r#"src="https://drive-d.mdview.example/my%20pics/b.png""#,
            r#"src="https://unc.mdview.example/server/share/c.png""#,
            r#"src="https://unc.mdview.example/server/share/c2.png""#,
            r#"src="%5Cserver%5Cshare%5Cx.png""#,
            r#"src="https://mdview.example/docs/img/d.png""#,
            r#"src="img/e.png""#,
            r#"href="https://drive-d.mdview.example/notes/f.md""#,
            r#"href="https://example.com/g""#,
            r#"<img src="D:\raw.png">"#,
        ] {
            assert!(html.contains(expected), "missing {expected} in {html}");
        }
        assert!(!html.contains("pipe.mdview") && !html.contains("mdview.example/./"), "a pipe never gets an MDView URL: {html}");
        let console = markdown_to_html(md);
        assert!(console.contains(r#"src="D:%5Cpics%5Ca.png""#) && !console.contains("mdview.example"), "console output is not rewritten: {console}");
    }
    #[test]
    fn test_remote_images_load_only_when_allowed() {
        let allowed = wrap("", false, RemoteImages::Allowed, &[]);
        assert!(allowed.contains("img-src https://mdview.example https://*.mdview.example data: https: http:;"));
        let ask = wrap("", false, RemoteImages::AskToLoad, &document_bars(&page(RemoteImages::AskToLoad, false)));
        assert!(ask.contains("img-src https://mdview.example https://*.mdview.example data:;"));
        let blocked = wrap("", false, RemoteImages::Blocked, &[]);
        assert!(blocked.contains("img-src https://mdview.example https://*.mdview.example data:;"));
    }

    fn page(remote_images: RemoteImages, ask_to_register: bool) -> PageBars<'static> {
        PageBars { remote_images, ask_to_register, ask_network_files: false, closed: &[] }
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
        let page = wrap("", false, RemoteImages::Blocked, &bars);
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
        let none = wrap("", false, RemoteImages::Blocked, &[]);
        assert!(none.contains("var bars = [];"), "a page without bars gets an empty list");
    }

    // Bar text is written into MDView's own script; it must never be able to
    // end that script or start another one.
    #[test]
    fn test_bar_text_cannot_leave_the_script() {
        let text = "</script><script>x()</script> & <!-- \"q\" \\ \u{2028}\u{2029}\n";
        let page = wrap("", false, RemoteImages::Blocked, &[bar("b", text, &[], BarShow::Now)]);
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
        let page_html = wrap(&without, false, RemoteImages::Blocked, &[download_gate_bar()]);
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
        let page_html = wrap("", false, RemoteImages::AskToLoad, &bars);
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
