//! The keys of each program (F1 help, #59) and MDView's right-click menu.
//! Shared by the standalone viewer and the TC plugin; pure, so it is tested.

use crate::markdown::script_string;

/// A group of keys in the help.
pub struct Group {
    pub title: &'static str,
    pub keys: Vec<(&'static str, &'static str)>,
}

/// The keys of this program, as the F1 help lists them.
pub fn shortcuts(is_plugin: bool) -> Vec<Group> {
    let mut reading = vec![
        ("Ctrl+Home / Ctrl+End", "Top / bottom of the document"),
        ("Ctrl+mouse wheel, Ctrl+Plus / Ctrl+Minus", "Zoom in / out"),
        ("Ctrl+0", "Zoom to 100%"),
    ];
    let mut find = if is_plugin {
        vec![
            ("F7 / Ctrl+F", "Find (Total Commander's search)"),
            ("F3 / Shift+F3", "Next / previous match"),
        ]
    } else {
        vec![("Ctrl+F", "Find")]
    };
    find.extend([("Ctrl+C", "Copy"), ("Ctrl+A", "Select all"), ("Ctrl+P", "Print")]);
    let mut window = Vec::new();
    if is_plugin {
        reading.push(("N / P", "Next / previous file"));
        reading.push(("1 - 7", "Lister view modes"));
        window.push(("Esc", "Close Lister"));
    } else {
        window.push(("Ctrl+O", "Open a file"));
        window.push(("Esc", "Close the window"));
    }
    window.extend([("Ctrl+,", "Settings"), ("F1", "These keyboard shortcuts"), ("Right-click", "Menu")]);
    vec![
        Group { title: "Reading", keys: reading },
        Group { title: "Find, copy and print", keys: find },
        Group { title: "Window", keys: window },
    ]
}

/// JSON for `PostWebMessageAsJson`: the page shows the keyboard help.
pub fn help_message(is_plugin: bool, dark: bool) -> String {
    let groups: Vec<String> = shortcuts(is_plugin)
        .iter()
        .map(|g| {
            let keys: Vec<String> = g
                .keys
                .iter()
                .map(|(k, a)| format!("[{},{}]", script_string(k), script_string(a)))
                .collect();
            format!(r#"{{"title":{},"keys":[{}]}}"#, script_string(g.title), keys.join(","))
        })
        .collect();
    format!(
        r#"{{"type":"showHelp","program":{},"dark":{dark},"groups":[{}]}}"#,
        script_string(if is_plugin { "Total Commander plugin" } else { "Standalone viewer" }),
        groups.join(",")
    )
}

/// What the right-click menu can do.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MenuCommand {
    OpenLink = 1,
    CopyLink,
    Copy,
    SelectAll,
    Find,
    Print,
    ZoomIn,
    ZoomOut,
    ZoomReset,
    Settings,
    Help,
}

impl MenuCommand {
    pub fn from_id(id: u32) -> Option<Self> {
        use MenuCommand::*;
        [OpenLink, CopyLink, Copy, SelectAll, Find, Print, ZoomIn, ZoomOut, ZoomReset, Settings, Help]
            .into_iter()
            .find(|c| *c as u32 == id)
    }
}

/// One line of the right-click menu; `None` is a separator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MenuItem {
    pub command: MenuCommand,
    /// The text, then a tab and the key, as Windows menus show it.
    pub label: String,
    pub enabled: bool,
}

/// The right-click menu at the place clicked.
pub fn menu_items(on_link: bool, has_selection: bool, zoom: u32, is_plugin: bool) -> Vec<Option<MenuItem>> {
    let item = |command, label: &str, enabled| Some(MenuItem { command, label: label.to_string(), enabled });
    let mut items = Vec::new();
    if on_link {
        items.push(item(MenuCommand::OpenLink, "&Open link", true));
        items.push(item(MenuCommand::CopyLink, "Copy &link address", true));
        items.push(None);
    }
    items.push(item(MenuCommand::Copy, "&Copy\tCtrl+C", has_selection));
    items.push(item(MenuCommand::SelectAll, "Select &all\tCtrl+A", true));
    items.push(None);
    items.push(item(MenuCommand::Find, if is_plugin { "&Find...\tF7" } else { "&Find...\tCtrl+F" }, true));
    items.push(item(MenuCommand::Print, "&Print...\tCtrl+P", true));
    items.push(None);
    items.push(item(MenuCommand::ZoomIn, "Zoom &in\tCtrl+Plus", true));
    items.push(item(MenuCommand::ZoomOut, "Zoom o&ut\tCtrl+Minus", true));
    items.push(item(MenuCommand::ZoomReset, &format!("&Reset zoom ({zoom}%)\tCtrl+0"), zoom != 100));
    items.push(None);
    items.push(item(MenuCommand::Settings, "&Settings...\tCtrl+,", true));
    items.push(item(MenuCommand::Help, "&Keyboard shortcuts\tF1", true));
    items
}

/// A key the page passes on to Total Commander's Lister (plugin): N / P
/// (next / previous file) and 1-9 (view modes). Anything else is refused.
#[allow(dead_code)] // plugin only
pub fn lister_key(key: &str) -> Option<u16> {
    match key {
        "N" | "n" => Some(b'N' as u16),
        "P" | "p" => Some(b'P' as u16),
        k if k.len() == 1 && matches!(k.as_bytes()[0], b'1'..=b'9') => Some(k.as_bytes()[0] as u16),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn keys(is_plugin: bool) -> Vec<&'static str> {
        shortcuts(is_plugin).into_iter().flat_map(|g| g.keys.into_iter().map(|(k, _)| k)).collect()
    }

    #[test]
    fn test_each_program_lists_its_own_keys() {
        let viewer = keys(false);
        let plugin = keys(true);
        for both in ["Ctrl+,", "F1", "Ctrl+0", "Ctrl+P", "Ctrl+C", "Ctrl+A", "Esc", "Right-click", "Ctrl+Home / Ctrl+End"] {
            assert!(viewer.contains(&both) && plugin.contains(&both), "{both}");
        }
        assert!(viewer.contains(&"Ctrl+O") && !plugin.contains(&"Ctrl+O"));
        assert!(plugin.contains(&"N / P") && plugin.contains(&"F3 / Shift+F3") && !viewer.contains(&"N / P"));
    }

    #[test]
    fn test_help_message_is_json_with_text_as_strings() {
        let json = help_message(true, true);
        assert!(json.starts_with(r#"{"type":"showHelp","program":"Total Commander plugin","dark":true,"groups":[{"title":"Reading","keys":[["#), "{json}");
        assert!(json.contains(r#"["N / P","Next / previous file"]"#));
        assert!(help_message(false, false).contains(r#""program":"Standalone viewer","dark":false"#));
    }

    #[test]
    fn test_menu_items_follow_the_place_clicked() {
        let labels = |items: &[Option<MenuItem>]| -> Vec<String> {
            items.iter().map(|i| i.as_ref().map_or("-".to_string(), |i| i.label.clone())).collect()
        };
        let plain = menu_items(false, false, 100, false);
        assert!(!labels(&plain).iter().any(|l| l.contains("link")), "link items only on a link");
        let copy = plain.iter().flatten().find(|i| i.command == MenuCommand::Copy).unwrap();
        assert!(!copy.enabled, "Copy is greyed without a selection");
        let reset = plain.iter().flatten().find(|i| i.command == MenuCommand::ZoomReset).unwrap();
        assert!(!reset.enabled && reset.label.contains("(100%)"));
        let link = menu_items(true, true, 125, true);
        assert_eq!(link[0].as_ref().unwrap().command, MenuCommand::OpenLink);
        assert_eq!(link[1].as_ref().unwrap().command, MenuCommand::CopyLink);
        assert!(link.iter().flatten().find(|i| i.command == MenuCommand::Copy).unwrap().enabled);
        assert!(labels(&link).contains(&"&Reset zoom (125%)\tCtrl+0".to_string()));
        assert!(labels(&link).contains(&"&Find...\tF7".to_string()), "plugin: TC's search");
        assert!(labels(&plain).contains(&"&Find...\tCtrl+F".to_string()));
        for item in link.iter().flatten() {
            assert_eq!(MenuCommand::from_id(item.command as u32), Some(item.command));
        }
        assert_eq!(MenuCommand::from_id(0), None);
    }

    #[test]
    fn test_only_lister_keys_are_passed_on() {
        assert_eq!(lister_key("N"), Some(0x4E));
        assert_eq!(lister_key("p"), Some(0x50));
        assert_eq!(lister_key("1"), Some(0x31));
        assert_eq!(lister_key("9"), Some(0x39));
        for no in ["0", "Q", "Escape", "F3", "", "12", "N\0"] {
            assert_eq!(lister_key(no), None, "{no:?}");
        }
    }
}
