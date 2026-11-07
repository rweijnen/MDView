//! Messages and questions shown inside the viewer instead of message boxes
//! (1.0.0). Shared by the standalone viewer and the TC plugin.
//!
//! The host sends `showMessage` with an id; the page confirms with
//! `messageShown` and answers with `messageChoice`. A question's action runs
//! only for the answer to a message MDView itself sent (matching id). If the
//! page doesn't confirm in time (it may be reloading), the host shows a
//! message box instead, so a message is never lost.

use std::collections::HashMap;

use crate::markdown::script_string;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Tone {
    #[allow(dead_code)] // standalone viewer only
    Info,
    Warning,
    Error,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Button {
    pub choice: &'static str,
    pub label: String,
}

/// A message or a question.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Message {
    pub tone: Tone,
    pub text: String,
    pub buttons: Vec<Button>,
    /// The safe answer: focused first, chosen by Esc, and the answer when
    /// the message is closed any other way.
    pub safe: &'static str,
}

pub const OK: &str = "ok";
pub const YES: &str = "yes";
pub const NO: &str = "no";

impl Message {
    pub fn error(text: impl Into<String>) -> Self {
        Message { tone: Tone::Error, text: text.into(), buttons: vec![Button { choice: OK, label: "OK".into() }], safe: OK }
    }

    pub fn warning(text: impl Into<String>) -> Self {
        Message { tone: Tone::Warning, text: text.into(), buttons: vec![Button { choice: OK, label: "OK".into() }], safe: OK }
    }

    /// A question: `yes_label` does it, Cancel (the safe answer) doesn't.
    pub fn confirm(tone: Tone, text: impl Into<String>, yes_label: &str) -> Self {
        Message {
            tone,
            text: text.into(),
            buttons: vec![Button { choice: YES, label: yes_label.into() }, Button { choice: NO, label: "Cancel".into() }],
            safe: NO,
        }
    }

    pub fn is_question(&self) -> bool {
        self.buttons.iter().any(|b| b.choice == YES)
    }
}

/// JSON for `PostWebMessageAsJson`: the page shows the message.
pub fn show_json(id: u64, message: &Message, dark: bool) -> String {
    let tone = match message.tone {
        Tone::Info => "info",
        Tone::Warning => "warning",
        Tone::Error => "error",
    };
    let buttons: Vec<String> = message
        .buttons
        .iter()
        .map(|b| format!(r#"{{"choice":{},"label":{}}}"#, script_string(b.choice), script_string(&b.label)))
        .collect();
    format!(
        r#"{{"type":"showMessage","id":"{id}","tone":"{tone}","dark":{dark},"text":{},"buttons":[{}],"safe":{}}}"#,
        script_string(&message.text),
        buttons.join(","),
        script_string(message.safe)
    )
}

/// Messages sent to a page and not answered yet, with the action to run on
/// "yes". The plugin keeps one per viewer window.
pub struct Pending<A> {
    next: u64,
    items: HashMap<u64, Item<A>>,
}

pub struct Item<A> {
    pub message: Message,
    pub action: Option<A>,
    pub shown: bool,
}

impl<A> Default for Pending<A> {
    fn default() -> Self {
        Pending { next: 1, items: HashMap::new() }
    }
}

impl<A> Pending<A> {
    pub fn add(&mut self, message: Message, action: Option<A>) -> u64 {
        let id = self.next;
        self.next += 1;
        self.items.insert(id, Item { message, action, shown: false });
        id
    }

    /// The page showed it: no message box needed.
    pub fn shown(&mut self, id: u64) {
        if let Some(item) = self.items.get_mut(&id) {
            item.shown = true;
        }
    }

    /// The page's answer. Returns the action to run when the answer was
    /// "yes" to a question MDView asked; `None` otherwise. The message is
    /// finished either way.
    pub fn answer(&mut self, id: u64, choice: &str) -> Option<A> {
        let item = self.items.remove(&id)?;
        if choice == YES && item.message.is_question() { item.action } else { None }
    }

    /// The page didn't confirm in time: the message, to show in a message box.
    pub fn not_shown(&mut self, id: u64) -> Option<Item<A>> {
        match self.items.get(&id) {
            Some(item) if !item.shown => self.items.remove(&id),
            _ => None,
        }
    }

    /// The page went away (navigation, closing): a shown message is gone,
    /// as if the safe answer was given. Unshown ones stay for `not_shown`.
    pub fn page_left(&mut self) {
        self.items.retain(|_, item| !item.shown);
    }
}

/// Parse the id the page sends back.
pub fn parse_id(text: &str) -> Option<u64> {
    text.parse().ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_show_json_is_json_with_text_as_strings() {
        let m = Message::confirm(Tone::Warning, "Open <x>?\n\"C:\\a.pdf\"", "Open");
        let json = show_json(7, &m, true);
        assert!(json.starts_with(r#"{"type":"showMessage","id":"7","tone":"warning","dark":true,"#), "{json}");
        assert!(json.contains(r#""text":"Open \u003cx\u003e?\n\"C:\\a.pdf\"""#), "{json}");
        assert!(json.contains(r#""buttons":[{"choice":"yes","label":"Open"},{"choice":"no","label":"Cancel"}],"safe":"no""#), "{json}");
        let e = show_json(1, &Message::error("x"), false);
        assert!(e.contains(r#""tone":"error","dark":false"#) && e.contains(r#"[{"choice":"ok","label":"OK"}],"safe":"ok""#));
    }

    // Only "yes" to a question MDView asked runs its action, and only once.
    #[test]
    fn test_only_yes_to_our_own_question_runs_the_action() {
        let mut p: Pending<&str> = Pending::default();
        let id = p.add(Message::confirm(Tone::Warning, "Run?", "Run"), Some("run"));
        assert_eq!(p.answer(id + 100, YES), None, "an id MDView never sent");
        assert_eq!(p.answer(id, YES), Some("run"));
        assert_eq!(p.answer(id, YES), None, "answered already");
        let id = p.add(Message::confirm(Tone::Warning, "Run?", "Run"), Some("run"));
        assert_eq!(p.answer(id, NO), None);
        let id = p.add(Message::confirm(Tone::Warning, "Run?", "Run"), Some("run"));
        assert_eq!(p.answer(id, "something"), None);
        // An error has no "yes": nothing runs even if the page says so.
        let id = p.add(Message::error("x"), Some("run"));
        assert_eq!(p.answer(id, YES), None);
    }

    // Not confirmed by the page: shown as a message box instead, once.
    #[test]
    fn test_unconfirmed_messages_fall_back_once() {
        let mut p: Pending<&str> = Pending::default();
        let a = p.add(Message::error("a"), None);
        let b = p.add(Message::error("b"), None);
        p.shown(b);
        assert_eq!(p.not_shown(a).map(|i| i.message.text), Some("a".to_string()));
        assert!(p.not_shown(a).is_none(), "only once");
        assert!(p.not_shown(b).is_none(), "shown on the page");
        p.page_left();
        assert_eq!(p.answer(b, OK), None, "a shown message is gone with its page");
        assert_eq!(parse_id("12"), Some(12));
        assert_eq!(parse_id("x"), None);
    }
}
