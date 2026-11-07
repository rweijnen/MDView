//! The settings panel (#50): what MDView sends the page to show, and checking
//! and saving what the page sends back. Shared by the standalone viewer and
//! the TC plugin; each saves its own settings (#34).

use std::collections::HashMap;

use crate::fonts::Font;
use crate::markdown::script_string;
use crate::settings::{self, AskSetting, FontSettings};

/// Everything the panel shows and saves.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSettings {
    pub fonts: FontSettings,
    pub remote_images: AskSetting,
    pub network_files: AskSetting,
    pub ask_downloads: bool,
    /// Standalone viewer only: offer to register MDView as .md viewer.
    pub ask_register: Option<bool>,
}

fn ask_name(setting: AskSetting) -> &'static str {
    match setting {
        AskSetting::Ask => "ask",
        AskSetting::Always => "always",
        AskSetting::Never => "never",
    }
}

fn parse_ask(value: &str) -> Option<AskSetting> {
    match value {
        "ask" => Some(AskSetting::Ask),
        "always" => Some(AskSetting::Always),
        "never" => Some(AskSetting::Never),
        _ => None,
    }
}

/// The settings now saved for this program, and notes on saved values that
/// are not valid.
pub fn current(is_plugin: bool) -> (PageSettings, Vec<String>) {
    let (fonts, notes) = settings::font_settings();
    let page = PageSettings {
        fonts,
        remote_images: settings::remote_images_setting(),
        network_files: settings::network_files_setting(),
        ask_downloads: !settings::open_downloaded_without_asking(),
        ask_register: (!is_plugin).then(|| !settings::dont_ask_register()),
    };
    (page, notes)
}

/// JSON for `PostWebMessageAsJson`: the page opens the panel with these.
pub fn show_message(page: &PageSettings, notes: &[String], fonts: &[Font], is_plugin: bool, dark: bool) -> String {
    let font_list: Vec<String> = fonts
        .iter()
        .map(|f| format!(r#"{{"name":{},"mono":{}}}"#, script_string(&f.name), f.monospace))
        .collect();
    let note_list: Vec<String> = notes.iter().map(|n| script_string(n)).collect();
    let register = match page.ask_register {
        Some(ask) => ask.to_string(),
        None => "null".to_string(),
    };
    format!(
        concat!(
            r#"{{"type":"showSettings","program":{},"dark":{},"#,
            r#""settings":{{"textFont":{},"textSize":{},"headingFont":{},"codeFont":{},"codeSize":{},"#,
            r#""remoteImages":"{}","networkFiles":"{}","askDownloads":{},"askRegister":{}}},"#,
            r#""fonts":[{}],"notes":[{}]}}"#
        ),
        script_string(if is_plugin { "Total Commander plugin" } else { "Standalone viewer" }),
        dark,
        script_string(&page.fonts.text_font),
        page.fonts.text_size,
        script_string(&page.fonts.heading_font),
        script_string(&page.fonts.code_font),
        page.fonts.code_size,
        ask_name(page.remote_images),
        ask_name(page.network_files),
        page.ask_downloads,
        register,
        font_list.join(","),
        note_list.join(","),
    )
}

pub fn saved_message() -> String {
    r#"{"type":"settingsSaved"}"#.to_string()
}

pub fn error_message(text: &str) -> String {
    format!(r#"{{"type":"settingsError","message":{}}}"#, script_string(text))
}

/// Check the page's `saveSettings` message. A font must be installed (or be
/// the one saved now, which may have been uninstalled since), sizes in range,
/// every field present and nothing else.
pub fn parse_save(
    msg: &HashMap<String, String>,
    is_plugin: bool,
    installed: &[Font],
    saved: &FontSettings,
) -> Result<PageSettings, String> {
    let mut expected = vec![
        "type", "textFont", "textSize", "headingFont", "codeFont", "codeSize", "remoteImages",
        "networkFiles", "askDownloads",
    ];
    if !is_plugin {
        expected.push("askRegister");
    }
    let mut problems = Vec::new();
    for key in msg.keys() {
        if !expected.contains(&key.as_str()) {
            problems.push(format!("unknown setting {key}"));
        }
    }
    let field = |key: &str| msg.get(key).map(String::as_str);
    let font = |key: &str, saved_now: &str, problems: &mut Vec<String>| -> String {
        match field(key) {
            Some("") => String::new(),
            Some(name) if name == saved_now || installed.iter().any(|f| f.name == name) => name.to_string(),
            Some(name) => {
                problems.push(format!("{name} is not an installed font"));
                String::new()
            }
            None => {
                problems.push(format!("{key} is missing"));
                String::new()
            }
        }
    };
    let size = |key: &str, zero_allowed: bool, problems: &mut Vec<String>| -> u32 {
        match field(key).and_then(|v| v.parse::<u32>().ok()) {
            Some(0) if zero_allowed => 0,
            Some(n) if settings::is_font_size(n) => n,
            _ => {
                problems.push(format!("{key} must be {}-{} px", settings::MIN_FONT_SIZE, settings::MAX_FONT_SIZE));
                0
            }
        }
    };
    let ask = |key: &str, problems: &mut Vec<String>| match field(key).and_then(parse_ask) {
        Some(a) => a,
        None => {
            problems.push(format!("{key} must be ask, always or never"));
            AskSetting::Ask
        }
    };
    let flag = |key: &str, problems: &mut Vec<String>| match field(key) {
        Some("true") => true,
        Some("false") => false,
        _ => {
            problems.push(format!("{key} must be true or false"));
            true
        }
    };
    let fonts = FontSettings {
        text_font: font("textFont", &saved.text_font, &mut problems),
        text_size: size("textSize", false, &mut problems),
        heading_font: font("headingFont", &saved.heading_font, &mut problems),
        code_font: font("codeFont", &saved.code_font, &mut problems),
        code_size: size("codeSize", true, &mut problems),
    };
    let page = PageSettings {
        fonts,
        remote_images: ask("remoteImages", &mut problems),
        network_files: ask("networkFiles", &mut problems),
        ask_downloads: flag("askDownloads", &mut problems),
        ask_register: (!is_plugin).then(|| flag("askRegister", &mut problems)),
    };
    if problems.is_empty() {
        Ok(page)
    } else {
        problems.sort();
        Err(format!("Nothing was saved. These settings are not valid: {}.", problems.join("; ")))
    }
}

/// Save every setting. Each is written on its own; if any fails, the message
/// names what was not saved and why, and says the others were saved.
pub fn save(page: &PageSettings) -> Result<(), String> {
    let mut failed: Vec<(&'static str, String)> = settings::set_font_settings(&page.fonts);
    let mut attempted = settings::FONT_VALUES;
    let mut check = |label: &'static str, result: Result<(), String>| {
        attempted += 1;
        if let Err(e) = result {
            failed.push((label, e));
        }
    };
    check("Remote images", settings::set_remote_images_setting(page.remote_images));
    check("Network files", settings::set_network_files_setting(page.network_files));
    check("Downloaded documents", settings::set_ask_before_downloaded(page.ask_downloads));
    if let Some(ask) = page.ask_register {
        check("Offer to register", settings::set_dont_ask_register(!ask));
    }
    failure_text(&failed, attempted)
}

fn failure_text(failed: &[(&'static str, String)], total: usize) -> Result<(), String> {
    if failed.is_empty() {
        return Ok(());
    }
    let labels: Vec<&str> = failed.iter().map(|(l, _)| *l).collect();
    let mut reasons: Vec<&str> = failed.iter().map(|(_, e)| e.as_str()).collect();
    reasons.dedup();
    let others = if failed.len() < total { "\nThe other settings were saved." } else { "" };
    Err(format!("Not saved: {}.\n{}{others}", labels.join(", "), reasons.join("\n")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn installed() -> Vec<Font> {
        ["Georgia", "Segoe UI", "Consolas"]
            .iter()
            .map(|n| Font { name: n.to_string(), monospace: *n == "Consolas" })
            .collect()
    }

    fn message(pairs: &[(&str, &str)]) -> HashMap<String, String> {
        pairs.iter().map(|(k, v)| (k.to_string(), v.to_string())).collect()
    }

    fn valid(extra: &[(&str, &str)]) -> HashMap<String, String> {
        let mut m = message(&[
            ("type", "saveSettings"), ("textFont", "Georgia"), ("textSize", "17"), ("headingFont", ""),
            ("codeFont", "Consolas"), ("codeSize", "0"), ("remoteImages", "always"),
            ("networkFiles", "never"), ("askDownloads", "false"),
        ]);
        m.extend(message(extra));
        m
    }

    #[test]
    fn test_a_valid_save_is_accepted() {
        let page = parse_save(&valid(&[("askRegister", "true")]), false, &installed(), &FontSettings::default()).unwrap();
        assert_eq!(page.fonts.text_font, "Georgia");
        assert_eq!(page.fonts.text_size, 17);
        assert_eq!(page.fonts.heading_font, "");
        assert_eq!(page.fonts.code_size, 0);
        assert_eq!(page.remote_images, AskSetting::Always);
        assert_eq!(page.network_files, AskSetting::Never);
        assert!(!page.ask_downloads);
        assert_eq!(page.ask_register, Some(true));
        // The plugin has no register setting.
        let page = parse_save(&valid(&[]), true, &installed(), &FontSettings::default()).unwrap();
        assert_eq!(page.ask_register, None);
    }

    // Anything the panel would never send is refused as a whole.
    #[test]
    fn test_invalid_saves_are_refused_and_nothing_is_saved() {
        let saved = FontSettings::default();
        for (key, value) in [
            ("textFont", "Not Installed"), ("textFont", "x\"; }"), ("textSize", "7"), ("textSize", "big"),
            ("codeSize", "49"), ("remoteImages", "maybe"), ("askDownloads", "yes"), ("color", "red"),
        ] {
            let err = parse_save(&valid(&[(key, value)]), true, &installed(), &saved).unwrap_err();
            assert!(err.starts_with("Nothing was saved."), "{key}={value}: {err}");
        }
        assert!(parse_save(&valid(&[("askRegister", "true")]), true, &installed(), &saved).is_err(), "plugin has no register setting");
        let mut missing = valid(&[]);
        missing.remove("codeFont");
        assert!(parse_save(&missing, true, &installed(), &saved).unwrap_err().contains("codeFont is missing"));
        assert!(parse_save(&valid(&[]), false, &installed(), &saved).is_err(), "viewer must send askRegister");
    }

    // A font saved earlier and uninstalled since can stay chosen.
    #[test]
    fn test_the_saved_font_may_stay_even_if_uninstalled() {
        let saved = FontSettings { text_font: "Old Font".into(), ..FontSettings::default() };
        let page = parse_save(&valid(&[("textFont", "Old Font")]), true, &installed(), &saved).unwrap();
        assert_eq!(page.fonts.text_font, "Old Font");
    }

    #[test]
    fn test_show_message_is_json_with_text_as_strings() {
        let page = PageSettings {
            fonts: FontSettings { text_font: "Georgia".into(), ..FontSettings::default() },
            remote_images: AskSetting::Ask,
            network_files: AskSetting::Always,
            ask_downloads: true,
            ask_register: None,
        };
        let json = show_message(&page, &["a <note>".to_string()], &installed(), true, true);
        assert!(json.starts_with(r#"{"type":"showSettings","program":"Total Commander plugin","dark":true,"#), "{json}");
        assert!(json.contains(r#""textFont":"Georgia","textSize":14,"headingFont":"","#));
        assert!(json.contains(r#""remoteImages":"ask","networkFiles":"always","askDownloads":true,"askRegister":null"#));
        assert!(json.contains(r#"{"name":"Consolas","mono":true}"#));
        assert!(json.contains(r#""notes":["a \u003cnote\u003e"]"#), "{json}");
        assert_eq!(error_message("a \"b\"\nc"), r#"{"type":"settingsError","message":"a \"b\"\nc"}"#);
    }

    // The message names what was not saved and why, as in the design.
    #[test]
    fn test_failure_text_names_what_failed() {
        assert_eq!(failure_text(&[], 9), Ok(()));
        let failed = [("Text size", "cannot write TextSize to C:\\x\\mdview.ini: Access is denied.".to_string()),
                      ("Code font", "cannot write TextSize to C:\\x\\mdview.ini: Access is denied.".to_string())];
        let text = failure_text(&failed, 9).unwrap_err();
        assert_eq!(text, "Not saved: Text size, Code font.\ncannot write TextSize to C:\\x\\mdview.ini: Access is denied.\nThe other settings were saved.");
        // Everything failed (the plugin saves 8 values): nothing else was saved.
        let all: Vec<(&'static str, String)> = (0..8).map(|_| ("x", "no settings folder".to_string())).collect();
        assert!(!failure_text(&all, 8).unwrap_err().contains("other settings were saved"));
        assert!(failure_text(&all[..7], 8).unwrap_err().contains("other settings were saved"));
    }
}
