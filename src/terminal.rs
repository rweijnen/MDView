//! Terminal markdown rendering with ANSI formatting and modern terminal features

use pulldown_cmark::{Event, Options, Parser, Tag, TagEnd, HeadingLevel, CodeBlockKind};
use std::env;

/// Terminal capabilities detected at runtime
#[derive(Debug, Clone)]
pub struct TerminalCaps {
    pub true_color: bool,
    pub hyperlinks: bool,
    pub unicode: bool,
    pub basic_ansi: bool,
}

impl TerminalCaps {
    /// Detect terminal capabilities from environment
    pub fn detect() -> Self {
        let is_windows_terminal = env::var("WT_SESSION").is_ok();
        let is_vscode = env::var("VSCODE_INJECTION").is_ok()
            || env::var("TERM_PROGRAM").ok().as_deref() == Some("vscode");
        let is_conemu = env::var("ConEmuPID").is_ok();
        let colorterm = env::var("COLORTERM").unwrap_or_default();
        let term = env::var("TERM").unwrap_or_default();

        // Windows Terminal supports everything
        if is_windows_terminal {
            return Self {
                true_color: true,
                hyperlinks: true,
                unicode: true,
                basic_ansi: true,
            };
        }

        // VS Code terminal
        if is_vscode {
            return Self {
                true_color: true,
                hyperlinks: true,
                unicode: true,
                basic_ansi: true,
            };
        }

        // ConEmu/Cmder
        if is_conemu {
            return Self {
                true_color: true,
                hyperlinks: true,
                unicode: true,
                basic_ansi: true,
            };
        }

        // Check for true color support
        let true_color = colorterm == "truecolor" || colorterm == "24bit"
            || term.contains("256color") || term.contains("truecolor")
            || cfg!(windows); // Windows Terminal and modern Windows consoles support true color

        // Hyperlinks (OSC 8) - only enable for known-good terminals
        // On legacy cmd.exe, show URL in parentheses so user can see/copy it
        let hyperlinks = term.contains("xterm")
            || term.contains("vte")
            || term.contains("kitty")
            || term.contains("iterm");

        // Unicode support - assume yes for most modern terminals
        let unicode = !term.is_empty() || cfg!(windows);

        // Basic ANSI - almost universal on modern systems
        // On Windows, assume ANSI support (Windows 10+ has it by default)
        let basic_ansi = !term.is_empty() || cfg!(windows);

        Self {
            true_color,
            hyperlinks,
            unicode,
            basic_ansi,
        }
    }

    /// Force basic mode (no colors, no unicode)
    pub fn basic() -> Self {
        Self {
            true_color: false,
            hyperlinks: false,
            unicode: false,
            basic_ansi: false,
        }
    }
}

/// ANSI escape codes
mod ansi {
    pub const RESET: &str = "\x1b[0m";
    pub const BOLD: &str = "\x1b[1m";
    pub const DIM: &str = "\x1b[2m";
    pub const ITALIC: &str = "\x1b[3m";
    pub const UNDERLINE: &str = "\x1b[4m";
    pub const STRIKETHROUGH: &str = "\x1b[9m";

    // Basic colors (works everywhere)
    pub const FG_RED: &str = "\x1b[31m";
    pub const FG_GREEN: &str = "\x1b[32m";
    pub const FG_YELLOW: &str = "\x1b[33m";
    pub const FG_BLUE: &str = "\x1b[34m";
    pub const FG_MAGENTA: &str = "\x1b[35m";
    pub const FG_CYAN: &str = "\x1b[36m";
    pub const FG_WHITE: &str = "\x1b[37m";
    pub const FG_GRAY: &str = "\x1b[90m";

    pub const BG_GRAY: &str = "\x1b[100m";

    /// True color foreground (24-bit RGB)
    pub fn fg_rgb(r: u8, g: u8, b: u8) -> String {
        format!("\x1b[38;2;{};{};{}m", r, g, b)
    }

    /// True color background (24-bit RGB)
    pub fn bg_rgb(r: u8, g: u8, b: u8) -> String {
        format!("\x1b[48;2;{};{};{}m", r, g, b)
    }

    /// OSC 8 hyperlink start
    pub fn hyperlink_start(url: &str) -> String {
        format!("\x1b]8;;{}\x1b\\", url)
    }

    /// OSC 8 hyperlink end
    pub const HYPERLINK_END: &str = "\x1b]8;;\x1b\\";
}

/// Unicode box drawing and symbols
mod unicode {
    pub const BULLET: char = '•';
    pub const CHECKBOX_UNCHECKED: &str = "☐";
    pub const CHECKBOX_CHECKED: &str = "☑";
    pub const QUOTE_BAR: &str = "│";
    pub const HORIZONTAL_LINE: &str = "─";

    // Table box drawing
    pub const TABLE_TOP_LEFT: char = '┌';
    pub const TABLE_TOP_RIGHT: char = '┐';
    pub const TABLE_BOTTOM_LEFT: char = '└';
    pub const TABLE_BOTTOM_RIGHT: char = '┘';
    pub const TABLE_HORIZONTAL: char = '─';
    pub const TABLE_VERTICAL: char = '│';
    pub const TABLE_CROSS: char = '┼';
    pub const TABLE_T_DOWN: char = '┬';
    pub const TABLE_T_UP: char = '┴';
    pub const TABLE_T_RIGHT: char = '├';
    pub const TABLE_T_LEFT: char = '┤';
}

/// Render markdown to terminal with ANSI formatting
pub fn render_to_terminal(markdown: &str, caps: &TerminalCaps) -> String {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);

    let parser = Parser::new_ext(markdown, options);
    let mut renderer = TerminalRenderer::new(caps.clone());

    for event in parser {
        renderer.process_event(event);
    }

    renderer.finish()
}

struct TerminalRenderer {
    caps: TerminalCaps,
    output: String,

    // State tracking
    in_heading: Option<HeadingLevel>,
    in_emphasis: bool,
    in_strong: bool,
    in_strikethrough: bool,
    in_code_block: bool,
    in_block_quote: u32,
    in_list: bool,
    list_index: Option<u64>,
    pending_link: Option<(String, String)>, // (url, title)
    link_text: String,

    // Table state
    in_table: bool,
    table_row: Vec<String>,
    table_rows: Vec<Vec<String>>,
    in_table_head: bool,
    current_cell: String,
}

impl TerminalRenderer {
    fn new(caps: TerminalCaps) -> Self {
        Self {
            caps,
            output: String::new(),
            in_heading: None,
            in_emphasis: false,
            in_strong: false,
            in_strikethrough: false,
            in_code_block: false,
            in_block_quote: 0,
            in_list: false,
            list_index: None,
            pending_link: None,
            link_text: String::new(),
            in_table: false,
            table_row: Vec::new(),
            table_rows: Vec::new(),
            in_table_head: false,
            current_cell: String::new(),
        }
    }

    fn process_event(&mut self, event: Event) {
        match event {
            Event::Start(tag) => self.start_tag(tag),
            Event::End(tag) => self.end_tag(tag),
            Event::Text(text) => self.text(&text),
            Event::Code(code) => self.inline_code(&code),
            Event::SoftBreak => self.soft_break(),
            Event::HardBreak => self.hard_break(),
            Event::Rule => self.horizontal_rule(),
            Event::TaskListMarker(checked) => self.task_list_marker(checked),
            _ => {}
        }
    }

    fn start_tag(&mut self, tag: Tag) {
        match tag {
            Tag::Heading { level, .. } => {
                self.in_heading = Some(level);
                self.output.push('\n');
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::BOLD);
                    let color = match level {
                        HeadingLevel::H1 => ansi::FG_MAGENTA,
                        HeadingLevel::H2 => ansi::FG_BLUE,
                        HeadingLevel::H3 => ansi::FG_CYAN,
                        _ => ansi::FG_GREEN,
                    };
                    self.output.push_str(color);
                }
                // No prefix - just colored/bold text
            }
            Tag::Paragraph => {
                if !self.output.is_empty() && !self.output.ends_with('\n') {
                    self.output.push('\n');
                }
                self.write_blockquote_prefix();
            }
            Tag::Emphasis => {
                self.in_emphasis = true;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::ITALIC);
                }
            }
            Tag::Strong => {
                self.in_strong = true;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::BOLD);
                }
            }
            Tag::Strikethrough => {
                self.in_strikethrough = true;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::STRIKETHROUGH);
                }
            }
            Tag::CodeBlock(kind) => {
                self.in_code_block = true;
                self.output.push('\n');
                if self.caps.basic_ansi {
                    if self.caps.true_color {
                        self.output.push_str(&ansi::bg_rgb(40, 44, 52));
                        self.output.push_str(&ansi::fg_rgb(171, 178, 191));
                    } else {
                        self.output.push_str(ansi::BG_GRAY);
                    }
                }
                // Show language if specified
                if let CodeBlockKind::Fenced(lang) = kind {
                    if !lang.is_empty() {
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::DIM);
                        }
                        self.output.push_str(&format!("  {}\n", lang));
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::RESET);
                            if self.caps.true_color {
                                self.output.push_str(&ansi::bg_rgb(40, 44, 52));
                                self.output.push_str(&ansi::fg_rgb(171, 178, 191));
                            } else {
                                self.output.push_str(ansi::BG_GRAY);
                            }
                        }
                    }
                }
            }
            Tag::BlockQuote(_) => {
                self.in_block_quote += 1;
                if !self.output.ends_with('\n') {
                    self.output.push('\n');
                }
            }
            Tag::List(start) => {
                self.in_list = true;
                self.list_index = start;
                if !self.output.ends_with('\n') {
                    self.output.push('\n');
                }
            }
            Tag::Item => {
                self.write_blockquote_prefix();
                if let Some(idx) = self.list_index.as_mut() {
                    self.output.push_str(&format!(" {}. ", idx));
                    *idx += 1;
                } else {
                    let bullet = if self.caps.unicode {
                        format!(" {} ", unicode::BULLET)
                    } else {
                        " * ".to_string()
                    };
                    self.output.push_str(&bullet);
                }
            }
            Tag::Link { dest_url, title, .. } => {
                self.pending_link = Some((dest_url.to_string(), title.to_string()));
                self.link_text.clear();
            }
            Tag::Image { dest_url, title, .. } => {
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::DIM);
                }
                self.output.push_str(&format!("[Image: {} ", dest_url));
                if !title.is_empty() {
                    self.output.push_str(&format!("\"{}\" ", title));
                }
                self.output.push(']');
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                }
            }
            Tag::Table(_) => {
                self.in_table = true;
                self.table_rows.clear();
                if !self.output.ends_with('\n') {
                    self.output.push('\n');
                }
            }
            Tag::TableHead => {
                self.in_table_head = true;
                self.table_row.clear();
            }
            Tag::TableRow => {
                self.table_row.clear();
            }
            Tag::TableCell => {
                self.current_cell.clear();
            }
            _ => {}
        }
    }

    fn end_tag(&mut self, tag: TagEnd) {
        match tag {
            TagEnd::Heading(_) => {
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                }
                self.output.push_str("\n\n");
                self.in_heading = None;
            }
            TagEnd::Paragraph => {
                self.output.push_str("\n\n");
            }
            TagEnd::Emphasis => {
                self.in_emphasis = false;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                    // Restore other active styles
                    self.restore_styles();
                }
            }
            TagEnd::Strong => {
                self.in_strong = false;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                    self.restore_styles();
                }
            }
            TagEnd::Strikethrough => {
                self.in_strikethrough = false;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                    self.restore_styles();
                }
            }
            TagEnd::CodeBlock => {
                self.in_code_block = false;
                if self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                }
                self.output.push('\n');
            }
            TagEnd::BlockQuote(_) => {
                self.in_block_quote = self.in_block_quote.saturating_sub(1);
            }
            TagEnd::List(_) => {
                self.in_list = false;
                self.list_index = None;
            }
            TagEnd::Item => {
                if !self.output.ends_with('\n') {
                    self.output.push('\n');
                }
            }
            TagEnd::Link => {
                if let Some((url, _title)) = self.pending_link.take() {
                    if self.caps.hyperlinks {
                        // OSC 8 clickable hyperlink
                        self.output.push_str(&ansi::hyperlink_start(&url));
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::FG_BLUE);
                            self.output.push_str(ansi::UNDERLINE);
                        }
                        self.output.push_str(&self.link_text);
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::RESET);
                        }
                        self.output.push_str(ansi::HYPERLINK_END);
                    } else {
                        // Fallback: show link text with URL in parentheses
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::FG_BLUE);
                            self.output.push_str(ansi::UNDERLINE);
                        }
                        self.output.push_str(&self.link_text);
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::RESET);
                        }
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::DIM);
                        }
                        self.output.push_str(&format!(" ({})", url));
                        if self.caps.basic_ansi {
                            self.output.push_str(ansi::RESET);
                        }
                    }
                }
                self.link_text.clear();
            }
            TagEnd::Table => {
                self.render_table();
                self.in_table = false;
            }
            TagEnd::TableHead => {
                self.in_table_head = false;
                self.table_rows.push(self.table_row.clone());
            }
            TagEnd::TableRow => {
                if !self.in_table_head {
                    self.table_rows.push(self.table_row.clone());
                }
            }
            TagEnd::TableCell => {
                self.table_row.push(self.current_cell.clone());
            }
            _ => {}
        }
    }

    fn text(&mut self, text: &str) {
        if self.pending_link.is_some() {
            self.link_text.push_str(text);
            return;
        }

        if self.in_table {
            self.current_cell.push_str(text);
            return;
        }

        if self.in_code_block {
            // Indent code block lines
            for line in text.lines() {
                self.output.push_str("  ");
                self.output.push_str(line);
                self.output.push('\n');
            }
        } else {
            self.output.push_str(text);
        }
    }

    fn inline_code(&mut self, code: &str) {
        if self.in_table {
            self.current_cell.push('`');
            self.current_cell.push_str(code);
            self.current_cell.push('`');
            return;
        }

        if self.caps.basic_ansi {
            if self.caps.true_color {
                self.output.push_str(&ansi::bg_rgb(40, 44, 52));
                self.output.push_str(&ansi::fg_rgb(230, 192, 123));
            } else {
                self.output.push_str(ansi::BG_GRAY);
                self.output.push_str(ansi::FG_YELLOW);
            }
        }
        self.output.push(' ');
        self.output.push_str(code);
        self.output.push(' ');
        if self.caps.basic_ansi {
            self.output.push_str(ansi::RESET);
        }
    }

    fn soft_break(&mut self) {
        if self.pending_link.is_some() {
            self.link_text.push(' ');
        } else if self.in_table {
            self.current_cell.push(' ');
        } else {
            self.output.push(' ');
        }
    }

    fn hard_break(&mut self) {
        self.output.push('\n');
        self.write_blockquote_prefix();
    }

    fn horizontal_rule(&mut self) {
        self.output.push('\n');
        if self.caps.basic_ansi {
            self.output.push_str(ansi::DIM);
        }
        let line = if self.caps.unicode {
            unicode::HORIZONTAL_LINE.repeat(40)
        } else {
            "-".repeat(40)
        };
        self.output.push_str(&line);
        if self.caps.basic_ansi {
            self.output.push_str(ansi::RESET);
        }
        self.output.push_str("\n\n");
    }

    fn task_list_marker(&mut self, checked: bool) {
        let marker = if self.caps.unicode {
            if checked { unicode::CHECKBOX_CHECKED } else { unicode::CHECKBOX_UNCHECKED }
        } else {
            if checked { "[x]" } else { "[ ]" }
        };
        if self.caps.basic_ansi {
            if checked {
                self.output.push_str(ansi::FG_GREEN);
            } else {
                self.output.push_str(ansi::DIM);
            }
        }
        self.output.push_str(marker);
        self.output.push(' ');
        if self.caps.basic_ansi {
            self.output.push_str(ansi::RESET);
        }
    }

    fn write_blockquote_prefix(&mut self) {
        if self.in_block_quote > 0 {
            if self.caps.basic_ansi {
                self.output.push_str(ansi::FG_GRAY);
            }
            for _ in 0..self.in_block_quote {
                let bar = if self.caps.unicode { unicode::QUOTE_BAR } else { "|" };
                self.output.push_str(bar);
                self.output.push(' ');
            }
            if self.caps.basic_ansi {
                self.output.push_str(ansi::RESET);
            }
        }
    }

    fn restore_styles(&mut self) {
        if self.in_strong {
            self.output.push_str(ansi::BOLD);
        }
        if self.in_emphasis {
            self.output.push_str(ansi::ITALIC);
        }
        if self.in_strikethrough {
            self.output.push_str(ansi::STRIKETHROUGH);
        }
    }

    fn render_table(&mut self) {
        if self.table_rows.is_empty() {
            return;
        }

        // Calculate column widths
        let col_count = self.table_rows.iter().map(|r| r.len()).max().unwrap_or(0);
        let mut col_widths: Vec<usize> = vec![0; col_count];

        for row in &self.table_rows {
            for (i, cell) in row.iter().enumerate() {
                if i < col_widths.len() {
                    col_widths[i] = col_widths[i].max(cell.chars().count());
                }
            }
        }

        // Ensure minimum width
        for w in &mut col_widths {
            *w = (*w).max(3);
        }

        let (h, v, tl, tr, bl, br, cross, td, tu, tleft, tright) = if self.caps.unicode {
            (unicode::TABLE_HORIZONTAL, unicode::TABLE_VERTICAL,
             unicode::TABLE_TOP_LEFT, unicode::TABLE_TOP_RIGHT,
             unicode::TABLE_BOTTOM_LEFT, unicode::TABLE_BOTTOM_RIGHT,
             unicode::TABLE_CROSS, unicode::TABLE_T_DOWN, unicode::TABLE_T_UP,
             unicode::TABLE_T_RIGHT, unicode::TABLE_T_LEFT)
        } else {
            ('-', '|', '+', '+', '+', '+', '+', '+', '+', '+', '+')
        };

        // Draw top border
        self.output.push(tl);
        for (i, &width) in col_widths.iter().enumerate() {
            self.output.push_str(&h.to_string().repeat(width + 2));
            if i < col_widths.len() - 1 {
                self.output.push(td);
            }
        }
        self.output.push(tr);
        self.output.push('\n');

        // Draw rows
        for (row_idx, row) in self.table_rows.iter().enumerate() {
            self.output.push(v);
            for (i, &width) in col_widths.iter().enumerate() {
                let cell = row.get(i).map(|s| s.as_str()).unwrap_or("");
                let padding = width.saturating_sub(cell.chars().count());
                self.output.push(' ');

                // Bold for header row
                if row_idx == 0 && self.caps.basic_ansi {
                    self.output.push_str(ansi::BOLD);
                }
                self.output.push_str(cell);
                if row_idx == 0 && self.caps.basic_ansi {
                    self.output.push_str(ansi::RESET);
                }

                self.output.push_str(&" ".repeat(padding + 1));
                self.output.push(v);
            }
            self.output.push('\n');

            // Draw separator after header
            if row_idx == 0 {
                self.output.push(tleft);
                for (i, &width) in col_widths.iter().enumerate() {
                    self.output.push_str(&h.to_string().repeat(width + 2));
                    if i < col_widths.len() - 1 {
                        self.output.push(cross);
                    }
                }
                self.output.push(tright);
                self.output.push('\n');
            }
        }

        // Draw bottom border
        self.output.push(bl);
        for (i, &width) in col_widths.iter().enumerate() {
            self.output.push_str(&h.to_string().repeat(width + 2));
            if i < col_widths.len() - 1 {
                self.output.push(tu);
            }
        }
        self.output.push(br);
        self.output.push_str("\n\n");
    }

    fn finish(mut self) -> String {
        // Trim trailing whitespace but keep one newline
        while self.output.ends_with("\n\n") {
            self.output.pop();
        }
        if !self.output.ends_with('\n') {
            self.output.push('\n');
        }
        self.output
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_basic_rendering() {
        let caps = TerminalCaps::basic();
        let output = render_to_terminal("# Hello\n\nWorld", &caps);
        assert!(output.contains("# Hello"));
        assert!(output.contains("World"));
    }

    #[test]
    fn test_detect_caps() {
        let caps = TerminalCaps::detect();
        // Should at least detect basic ANSI on most systems
        assert!(caps.basic_ansi || cfg!(not(windows)));
    }
}
