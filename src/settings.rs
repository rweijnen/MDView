//! Settings of the standalone viewer and the Total Commander plugin.
//!
//! Each program keeps its own settings, and answers to bars and prompts are
//! stored where they were given (#34):
//! - standalone viewer: `HKCU\Software\MDView`
//! - TC plugin: `mdview.ini` in TC's ini folder (next to wincmd.ini), as
//!   reported by TC through `ListSetDefaultParams` (WLX SDK)

use std::path::{Path, PathBuf};
use std::sync::OnceLock;
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};

use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, KEY_READ, KEY_WRITE, REG_DWORD, REG_OPTION_NON_VOLATILE, REG_SZ,
    REG_VALUE_TYPE, RegCloseKey, RegCreateKeyExW, RegOpenKeyExW, RegQueryValueExW,
    RegSetValueExW,
};
use windows::core::PCWSTR;

/// The product name, defined once. A macro so that compile-time constants
/// (registry keys, ProgID, message texts) can be built from it with `concat!`.
macro_rules! app_name {
    () => {
        "MDView"
    };
}
pub(crate) use app_name;

pub const APP_NAME: &str = app_name!();
pub const REGISTRY_KEY: &str = concat!("Software\\", app_name!());

const REMOTE_IMAGES: &str = "RemoteImages";

/// The plugin's own settings file, in TC's ini folder. Not TC's shared
/// lsplugin.ini: that file is shared with other plugins (WLX SDK).
#[allow(dead_code)] // plugin only; settings.rs is compiled into both binaries
pub const PLUGIN_INI_NAME: &str = concat!(app_name!(), ".ini");

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Where this program keeps its settings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Store {
    /// Standalone viewer: `HKCU\Software\MDView`.
    Registry,
    /// TC plugin: `mdview.ini` in TC's ini folder.
    #[allow(dead_code)] // constructed by the plugin only
    Ini(PathBuf),
    /// TC plugin without a usable ini folder from TC: nothing can be stored.
    Unavailable(String),
}

static STORE: OnceLock<Store> = OnceLock::new();

/// This program's settings store. The standalone viewer always uses the
/// registry; the plugin uses what `ListSetDefaultParams` set, and nothing if TC
/// never reported its ini folder (MDView only supports TC as plugin host).
fn store() -> &'static Store {
    STORE.get_or_init(|| {
        if crate::IS_PLUGIN {
            Store::Unavailable("Total Commander did not report its settings folder".to_string())
        } else {
            Store::Registry
        }
    })
}

/// The plugin's store for TC's ini folder (the folder of `DefaultIniName`).
#[allow(dead_code)] // plugin only
pub fn ini_store_for(ini_folder: &Path) -> Store {
    if ini_folder.as_os_str().is_empty() {
        Store::Unavailable("Total Commander reported an empty settings folder".to_string())
    } else if !ini_folder.is_dir() {
        Store::Unavailable(format!(
            "Total Commander's settings folder {} was not found",
            ini_folder.display()
        ))
    } else {
        Store::Ini(ini_folder.join(PLUGIN_INI_NAME.to_ascii_lowercase()))
    }
}

/// Set the plugin's settings store; called once, from ListSetDefaultParams.
/// Returns the store now in use (TC loads the plugin once per process).
#[allow(dead_code)] // plugin only
pub fn use_store(store: Store) -> &'static Store {
    let _ = STORE.set(store);
    self::store()
}

/// Read a numeric setting; `None` if it was never saved or cannot be read.
pub fn read_dword(name: &str) -> Option<u32> {
    read_dword_in(store(), name)
}

/// Save a numeric setting.
pub fn write_dword(name: &str, value: u32) -> Result<(), String> {
    write_dword_in(store(), name, value)
}

fn read_dword_in(store: &Store, name: &str) -> Option<u32> {
    match store {
        Store::Registry => registry_read_dword(name),
        Store::Ini(path) => ini_read_dword(path, name),
        Store::Unavailable(_) => None,
    }
}

fn write_dword_in(store: &Store, name: &str, value: u32) -> Result<(), String> {
    match store {
        Store::Registry => registry_write_dword(name, value),
        Store::Ini(path) => ini_write_dword(path, name, value),
        Store::Unavailable(reason) => Err(reason.clone()),
    }
}

/// `[MDView] name=value` in the plugin's ini file.
fn ini_read_dword(path: &Path, name: &str) -> Option<u32> {
    use windows::Win32::System::WindowsProgramming::GetPrivateProfileStringW;
    let section = wide(APP_NAME);
    let key = wide(name);
    let file = wide(&path.to_string_lossy());
    let empty = wide("");
    let mut buffer = [0u16; 32];
    let len = unsafe {
        GetPrivateProfileStringW(
            PCWSTR(section.as_ptr()),
            PCWSTR(key.as_ptr()),
            PCWSTR(empty.as_ptr()),
            Some(&mut buffer),
            PCWSTR(file.as_ptr()),
        )
    } as usize;
    String::from_utf16_lossy(&buffer[..len]).trim().parse().ok()
}

fn ini_write_dword(path: &Path, name: &str, value: u32) -> Result<(), String> {
    use windows::Win32::System::WindowsProgramming::WritePrivateProfileStringW;
    let section = wide(APP_NAME);
    let key = wide(name);
    let text = wide(&value.to_string());
    let file = wide(&path.to_string_lossy());
    unsafe {
        WritePrivateProfileStringW(
            PCWSTR(section.as_ptr()),
            PCWSTR(key.as_ptr()),
            PCWSTR(text.as_ptr()),
            PCWSTR(file.as_ptr()),
        )
    }
    .map_err(|e| format!("cannot write {name} to {}: {}", path.display(), e.message()))
}

/// Read a DWORD value; `None` if the key or value does not exist.
fn registry_read_dword(name: &str) -> Option<u32> {
    let subkey = wide(REGISTRY_KEY);
    let name = wide(name);
    unsafe {
        let mut hkey = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, PCWSTR(subkey.as_ptr()), Some(0), KEY_READ, &mut hkey)
            .is_err()
        {
            return None;
        }
        let mut value: u32 = 0;
        let mut size: u32 = 4;
        let mut data_type = REG_VALUE_TYPE::default();
        let result = RegQueryValueExW(
            hkey,
            PCWSTR(name.as_ptr()),
            None,
            Some(&mut data_type),
            Some(&mut value as *mut u32 as *mut u8),
            Some(&mut size),
        );
        let _ = RegCloseKey(hkey);
        (result.is_ok() && data_type == REG_DWORD && size == 4).then_some(value)
    }
}

/// Write a DWORD value, creating the key if needed.
fn registry_write_dword(name: &str, value: u32) -> Result<(), String> {
    let subkey = wide(REGISTRY_KEY);
    let name_wide = wide(name);
    unsafe {
        let mut hkey = HKEY::default();
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut hkey,
            None,
        )
        .ok()
        .map_err(|e| format!("cannot open HKCU\\{REGISTRY_KEY}: {}", e.message()))?;
        let result = RegSetValueExW(
            hkey,
            PCWSTR(name_wide.as_ptr()),
            Some(0),
            REG_DWORD,
            Some(&value.to_le_bytes()),
        );
        let _ = RegCloseKey(hkey);
        result
            .ok()
            .map_err(|e| format!("cannot write {name}: {}", e.message()))
    }
}

/// Whether something a document asks for is loaded: remote (http/https)
/// images, or files from a network location. Stored as 0/1/2.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AskSetting {
    /// Default: blocked, with a bar offering to load it.
    Ask,
    /// Always load it.
    Always,
    /// Never load it and don't ask.
    Never,
}

impl AskSetting {
    fn from_dword(v: u32) -> Self {
        match v {
            1 => Self::Always,
            2 => Self::Never,
            _ => Self::Ask,
        }
    }

    fn to_dword(self) -> u32 {
        match self {
            Self::Ask => 0,
            Self::Always => 1,
            Self::Never => 2,
        }
    }
}

pub fn remote_images_setting() -> AskSetting {
    read_dword(REMOTE_IMAGES)
        .map(AskSetting::from_dword)
        .unwrap_or(AskSetting::Ask)
}

pub fn set_remote_images_setting(setting: AskSetting) -> Result<(), String> {
    write_dword(REMOTE_IMAGES, setting.to_dword())
}

/// Files from another share or a network drive, named in a document. Separate
/// from remote images: a share can also receive the user's Windows login.
const NETWORK_FILES: &str = "NetworkFiles";

pub fn network_files_setting() -> AskSetting {
    read_dword(NETWORK_FILES).map(AskSetting::from_dword).unwrap_or(AskSetting::Ask)
}

pub fn set_network_files_setting(setting: AskSetting) -> Result<(), String> {
    write_dword(NETWORK_FILES, setting.to_dword())
}

/// "Don't ask again" on the page shown instead of a downloaded document (#24).
const OPEN_DOWNLOADED_WITHOUT_ASKING: &str = "OpenDownloadedWithoutAsking";

pub fn open_downloaded_without_asking() -> bool {
    read_dword(OPEN_DOWNLOADED_WITHOUT_ASKING) == Some(1)
}

pub fn set_open_downloaded_without_asking() -> Result<(), String> {
    write_dword(OPEN_DOWNLOADED_WITHOUT_ASKING, 1)
}

/// From the settings page (#50): `true` asks again before showing a
/// downloaded document.
pub fn set_ask_before_downloaded(ask: bool) -> Result<(), String> {
    write_dword(OPEN_DOWNLOADED_WITHOUT_ASKING, if ask { 0 } else { 1 })
}

/// "Don't ask again" in the registration bar (#37). Standalone viewer only.
const DONT_ASK_REGISTER: &str = "DontAskRegister";

#[allow(dead_code)] // standalone viewer only
pub fn dont_ask_register() -> bool {
    read_dword(DONT_ASK_REGISTER) == Some(1)
}

#[allow(dead_code)] // standalone viewer only
pub fn set_dont_ask_register(dont_ask: bool) -> Result<(), String> {
    write_dword(DONT_ASK_REGISTER, if dont_ask { 1 } else { 0 })
}

/// Read a text setting; `None` if it was never saved or cannot be read.
fn read_string_in(store: &Store, name: &str) -> Option<String> {
    match store {
        Store::Registry => registry_read_string(name),
        Store::Ini(path) => ini_read_string(path, name),
        Store::Unavailable(_) => None,
    }
}

fn write_string_in(store: &Store, name: &str, value: &str) -> Result<(), String> {
    match store {
        Store::Registry => registry_write_string(name, value),
        Store::Ini(path) => ini_write_string(path, name, value),
        Store::Unavailable(reason) => Err(reason.clone()),
    }
}

fn ini_read_string(path: &Path, name: &str) -> Option<String> {
    use windows::Win32::System::WindowsProgramming::GetPrivateProfileStringW;
    let section = wide(APP_NAME);
    let key = wide(name);
    let file = wide(&path.to_string_lossy());
    // A value that was never written reads as this marker, not as "".
    let missing = wide("\u{1}");
    let mut buffer = [0u16; 512];
    let len = unsafe {
        GetPrivateProfileStringW(
            PCWSTR(section.as_ptr()),
            PCWSTR(key.as_ptr()),
            PCWSTR(missing.as_ptr()),
            Some(&mut buffer),
            PCWSTR(file.as_ptr()),
        )
    } as usize;
    let value = String::from_utf16_lossy(&buffer[..len]);
    (value != "\u{1}").then_some(value)
}

fn ini_write_string(path: &Path, name: &str, value: &str) -> Result<(), String> {
    use windows::Win32::System::WindowsProgramming::WritePrivateProfileStringW;
    let section = wide(APP_NAME);
    let key = wide(name);
    let text = wide(value);
    let file = wide(&path.to_string_lossy());
    unsafe {
        WritePrivateProfileStringW(
            PCWSTR(section.as_ptr()),
            PCWSTR(key.as_ptr()),
            PCWSTR(text.as_ptr()),
            PCWSTR(file.as_ptr()),
        )
    }
    .map_err(|e| format!("cannot write {name} to {}: {}", path.display(), e.message()))
}

fn registry_read_string(name: &str) -> Option<String> {
    let subkey = wide(REGISTRY_KEY);
    let name = wide(name);
    unsafe {
        let mut hkey = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, PCWSTR(subkey.as_ptr()), Some(0), KEY_READ, &mut hkey)
            .is_err()
        {
            return None;
        }
        let mut buffer = [0u16; 512];
        let mut size = (buffer.len() * 2) as u32;
        let mut data_type = REG_VALUE_TYPE::default();
        let result = RegQueryValueExW(
            hkey,
            PCWSTR(name.as_ptr()),
            None,
            Some(&mut data_type),
            Some(buffer.as_mut_ptr() as *mut u8),
            Some(&mut size),
        );
        let _ = RegCloseKey(hkey);
        if result.is_err() || data_type != REG_SZ {
            return None;
        }
        let units = (size as usize / 2).min(buffer.len());
        let text = &buffer[..units];
        let end = text.iter().position(|&c| c == 0).unwrap_or(text.len());
        Some(String::from_utf16_lossy(&text[..end]))
    }
}

fn registry_write_string(name: &str, value: &str) -> Result<(), String> {
    let subkey = wide(REGISTRY_KEY);
    let name_wide = wide(name);
    let data: Vec<u8> = wide(value).iter().flat_map(|c| c.to_le_bytes()).collect();
    unsafe {
        let mut hkey = HKEY::default();
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            PCWSTR(subkey.as_ptr()),
            Some(0),
            None,
            REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            None,
            &mut hkey,
            None,
        )
        .ok()
        .map_err(|e| format!("cannot open HKCU\\{REGISTRY_KEY}: {}", e.message()))?;
        let result = RegSetValueExW(hkey, PCWSTR(name_wide.as_ptr()), Some(0), REG_SZ, Some(&data));
        let _ = RegCloseKey(hkey);
        result.ok().map_err(|e| format!("cannot write {name}: {}", e.message()))
    }
}

/// Fonts for the document (#50). An empty font name means MDView's own
/// default; a code size of 0 means 85% of the text size.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FontSettings {
    pub text_font: String,
    pub text_size: u32,
    /// Empty: the same font as the text.
    pub heading_font: String,
    pub code_font: String,
    pub code_size: u32,
}

pub const DEFAULT_TEXT_SIZE: u32 = 14;
pub const MIN_FONT_SIZE: u32 = 8;
pub const MAX_FONT_SIZE: u32 = 48;

impl Default for FontSettings {
    fn default() -> Self {
        FontSettings {
            text_font: String::new(),
            text_size: DEFAULT_TEXT_SIZE,
            heading_font: String::new(),
            code_font: String::new(),
            code_size: 0,
        }
    }
}

const TEXT_FONT: &str = "TextFont";
const TEXT_SIZE: &str = "TextSize";
const HEADING_FONT: &str = "HeadingFont";
const CODE_FONT: &str = "CodeFont";
const CODE_SIZE: &str = "CodeSize";

/// A font family name as Windows reports it: letters, digits, spaces and
/// `-_.&'()`, at most 64 characters. Nothing else can reach the page's CSS.
pub fn is_font_name(name: &str) -> bool {
    !name.is_empty()
        && name.chars().count() <= 64
        && name.trim() == name
        && name.chars().all(|c| c.is_alphanumeric() || " -_.&'()".contains(c))
}

pub fn is_font_size(size: u32) -> bool {
    (MIN_FONT_SIZE..=MAX_FONT_SIZE).contains(&size)
}

/// The saved fonts, and a note for each saved value that is not valid (it is
/// not used; the settings page shows the notes).
pub fn font_settings() -> (FontSettings, Vec<String>) {
    font_settings_in(store())
}

fn font_settings_in(store: &Store) -> (FontSettings, Vec<String>) {
    let mut fonts = FontSettings::default();
    let mut notes = Vec::new();
    let mut font = |key: &str, label: &str, target: &mut String| {
        if let Some(name) = read_string_in(store, key) {
            if name.is_empty() || is_font_name(&name) {
                *target = name;
            } else {
                notes.push(format!("The saved {label} font is not a valid font name and is not used."));
            }
        }
    };
    font(TEXT_FONT, "text", &mut fonts.text_font);
    font(HEADING_FONT, "heading", &mut fonts.heading_font);
    font(CODE_FONT, "code", &mut fonts.code_font);
    if let Some(size) = read_dword_in(store, TEXT_SIZE) {
        if is_font_size(size) {
            fonts.text_size = size;
        } else {
            notes.push(format!("The saved text size {size} is outside {MIN_FONT_SIZE}-{MAX_FONT_SIZE} px and is not used."));
        }
    }
    if let Some(size) = read_dword_in(store, CODE_SIZE) {
        if size == 0 || is_font_size(size) {
            fonts.code_size = size;
        } else {
            notes.push(format!("The saved code size {size} is outside {MIN_FONT_SIZE}-{MAX_FONT_SIZE} px and is not used."));
        }
    }
    (fonts, notes)
}

/// Save the fonts. Each value is written on its own; the result names every
/// one that could not be saved, with the reason.
pub fn set_font_settings(fonts: &FontSettings) -> Vec<(&'static str, String)> {
    set_font_settings_in(store(), fonts)
}

/// How many values `set_font_settings` writes.
pub const FONT_VALUES: usize = 5;

fn set_font_settings_in(store: &Store, fonts: &FontSettings) -> Vec<(&'static str, String)> {
    let results = [
        ("Text font", write_string_in(store, TEXT_FONT, &fonts.text_font)),
        ("Text size", write_dword_in(store, TEXT_SIZE, fonts.text_size)),
        ("Heading font", write_string_in(store, HEADING_FONT, &fonts.heading_font)),
        ("Code font", write_string_in(store, CODE_FONT, &fonts.code_font)),
        ("Code size", write_dword_in(store, CODE_SIZE, fonts.code_size)),
    ];
    results.into_iter().filter_map(|(label, r)| r.err().map(|e| (label, e))).collect()
}

/// Zoom of MDView's pages in percent, one of `markdown::ZOOM_STEPS` (#16).
const ZOOM: &str = "Zoom";

/// The zoom chosen in this run (0: none yet). Used for the next pages even if
/// saving it failed.
static SESSION_ZOOM: AtomicU32 = AtomicU32::new(0);
/// Saving the zoom failed before in this run; the user was told once.
static ZOOM_SAVE_FAILED: AtomicBool = AtomicBool::new(false);

fn zoom_in(store: &Store) -> u32 {
    read_dword_in(store, ZOOM)
        .filter(|&p| crate::markdown::is_zoom_step(p))
        .unwrap_or(crate::markdown::DEFAULT_ZOOM)
}

/// The zoom for the next page.
pub fn zoom() -> u32 {
    match SESSION_ZOOM.load(Ordering::Relaxed) {
        0 => zoom_in(store()),
        p => p,
    }
}

/// Use `percent` from now on and save it. Returns why saving failed, but only
/// the first time in this run, so zooming doesn't bring up a message per step.
pub fn set_zoom(percent: u32) -> Option<String> {
    SESSION_ZOOM.store(percent, Ordering::Relaxed);
    save_zoom_in(store(), percent, &ZOOM_SAVE_FAILED)
}

fn save_zoom_in(store: &Store, percent: u32, failed_before: &AtomicBool) -> Option<String> {
    match write_dword_in(store, ZOOM, percent) {
        Ok(()) => None,
        Err(e) => (!failed_before.swap(true, Ordering::Relaxed)).then_some(e),
    }
}

/// WebView2 data folder (cache, cookies; not settings) for this process.
///
/// A WebView2 data folder can only be used by one host program at a time:
/// with MDView open in Total Commander's Lister, a standalone viewer using the
/// same folder fails to load (and the other way round, #33). So each host
/// program gets its own folder, and elevated and normal processes are kept
/// apart, so files created elevated never have to be written by a normal
/// process. Lives in %LOCALAPPDATA%: local, per user, not roaming.
pub fn webview2_data_folder() -> Result<std::path::PathBuf, String> {
    let local_app_data = local_app_data_folder()?;
    Ok(local_app_data
        .join(APP_NAME)
        .join("WebView2")
        .join(data_folder_name(&host_program_name()?, is_elevated()?)))
}

/// `<host program>-<elevated|normal>`, using only characters that are safe
/// in a folder name.
fn data_folder_name(host_program: &str, elevated: bool) -> String {
    let host: String = host_program
        .to_lowercase()
        .chars()
        .map(|c| if c.is_ascii_alphanumeric() || c == '-' || c == '_' { c } else { '_' })
        .collect();
    format!("{}-{}", host, if elevated { "elevated" } else { "normal" })
}

/// The program MDView runs in: `mdview` for the standalone viewer, the file
/// manager (e.g. `totalcmd64`) for the plugin.
fn host_program_name() -> Result<String, String> {
    use windows::Win32::System::LibraryLoader::GetModuleFileNameW;
    let mut buffer = vec![0u16; 32768];
    let len = unsafe { GetModuleFileNameW(None, &mut buffer) } as usize;
    if len == 0 || len >= buffer.len() {
        return Err("cannot determine the host program".to_string());
    }
    let path = std::path::PathBuf::from(String::from_utf16_lossy(&buffer[..len]));
    path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .filter(|s| !s.is_empty())
        .ok_or_else(|| format!("cannot determine the host program from {}", path.display()))
}

fn is_elevated() -> Result<bool, String> {
    use windows::Win32::Foundation::{CloseHandle, HANDLE};
    use windows::Win32::Security::{GetTokenInformation, TOKEN_ELEVATION, TOKEN_QUERY, TokenElevation};
    use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};
    unsafe {
        let mut token = HANDLE::default();
        OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &mut token)
            .map_err(|e| format!("cannot open the process token: {}", e.message()))?;
        let mut elevation = TOKEN_ELEVATION::default();
        let mut size = 0u32;
        let result = GetTokenInformation(
            token,
            TokenElevation,
            Some(&mut elevation as *mut _ as *mut core::ffi::c_void),
            std::mem::size_of::<TOKEN_ELEVATION>() as u32,
            &mut size,
        );
        let _ = CloseHandle(token);
        result.map_err(|e| format!("cannot read the process elevation: {}", e.message()))?;
        Ok(elevation.TokenIsElevated != 0)
    }
}

fn local_app_data_folder() -> Result<std::path::PathBuf, String> {
    use windows::Win32::System::Com::CoTaskMemFree;
    use windows::Win32::UI::Shell::{FOLDERID_LocalAppData, KF_FLAG_DEFAULT, SHGetKnownFolderPath};
    unsafe {
        let path = SHGetKnownFolderPath(&FOLDERID_LocalAppData, KF_FLAG_DEFAULT, None)
            .map_err(|e| format!("cannot find the local application data folder: {}", e.message()))?;
        let result = path.to_string();
        CoTaskMemFree(Some(path.0 as *const _));
        result
            .map(std::path::PathBuf::from)
            .map_err(|_| "the local application data folder path is not valid".to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_data_folder_name_per_host_and_elevation() {
        assert_eq!(data_folder_name("mdview", false), "mdview-normal");
        assert_eq!(data_folder_name("TOTALCMD64", true), "totalcmd64-elevated");
        // Different programs and elevation never share a folder.
        assert_ne!(data_folder_name("mdview", false), data_folder_name("TOTALCMD64", false));
        assert_ne!(data_folder_name("mdview", false), data_folder_name("mdview", true));
        // Only safe characters end up in the folder name.
        assert_eq!(data_folder_name("Double Commander", false), "double_commander-normal");
        assert_eq!(data_folder_name("a.b\\c:d", true), "a_b_c_d-elevated");
    }

    #[test]
    fn test_data_folder_is_in_local_app_data_for_this_program() {
        let folder = webview2_data_folder().expect("data folder");
        let local = local_app_data_folder().expect("local app data");
        assert!(folder.starts_with(local.join(APP_NAME).join("WebView2")), "{}", folder.display());
        let name = folder.file_name().unwrap().to_string_lossy().into_owned();
        assert!(name.ends_with("-normal") || name.ends_with("-elevated"), "{name}");
    }

    /// A fresh folder for one test, removed afterwards.
    struct TempFolder(PathBuf);
    impl TempFolder {
        fn new(name: &str) -> Self {
            let dir = std::env::temp_dir().join(format!(
                "mdview-settings-test-{name}-{}",
                std::process::id()
            ));
            let _ = std::fs::remove_dir_all(&dir);
            std::fs::create_dir_all(&dir).unwrap();
            TempFolder(dir)
        }
    }
    impl Drop for TempFolder {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }

    // #34: the plugin keeps its settings in its own mdview.ini in TC's ini folder.
    #[test]
    fn test_ini_store_round_trips_settings() {
        let dir = TempFolder::new("roundtrip");
        let store = ini_store_for(&dir.0);
        assert_eq!(store, Store::Ini(dir.0.join("mdview.ini")));
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), None, "nothing saved yet");
        write_dword_in(&store, REMOTE_IMAGES, 1).unwrap();
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), Some(1));
        write_dword_in(&store, REMOTE_IMAGES, 2).unwrap();
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), Some(2));
        let text = std::fs::read_to_string(dir.0.join("mdview.ini")).unwrap();
        assert!(text.contains("[MDView]") && text.contains("RemoteImages=2"), "{text:?}");
    }

    #[test]
    fn test_ini_store_ignores_other_sections_and_bad_values() {
        let dir = TempFolder::new("sections");
        let path = dir.0.join("mdview.ini");
        std::fs::write(&path, "[Other]\r\nRemoteImages=1\r\n[MDView]\r\nRemoteImages=abc\r\n").unwrap();
        let store = Store::Ini(path);
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), None);
    }

    #[test]
    fn test_unusable_ini_folder_cannot_store() {
        let missing = std::env::temp_dir().join("mdview-settings-test-does-not-exist-34");
        let store = ini_store_for(&missing);
        assert!(matches!(&store, Store::Unavailable(reason) if reason.contains("was not found")), "{store:?}");
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), None);
        let err = write_dword_in(&store, REMOTE_IMAGES, 1).unwrap_err();
        assert!(err.contains("was not found"), "{err}");
        assert!(matches!(ini_store_for(Path::new("")), Store::Unavailable(_)));
    }

    #[test]
    fn test_remote_images_setting_round_trips_and_defaults_to_ask() {
        for s in [AskSetting::Ask, AskSetting::Always, AskSetting::Never] {
            assert_eq!(AskSetting::from_dword(s.to_dword()), s);
        }
        assert_eq!(AskSetting::from_dword(99), AskSetting::Ask);
    }

    // Network files have their own value: "Always load" for remote images
    // must not also send requests to a share.
    #[test]
    fn test_network_files_are_stored_apart_from_remote_images() {
        assert_ne!(NETWORK_FILES, REMOTE_IMAGES);
        let dir = TempFolder::new("network");
        let store = ini_store_for(&dir.0);
        write_dword_in(&store, REMOTE_IMAGES, AskSetting::Always.to_dword()).unwrap();
        assert_eq!(read_dword_in(&store, NETWORK_FILES), None, "remote images do not set network files");
        write_dword_in(&store, NETWORK_FILES, AskSetting::Never.to_dword()).unwrap();
        assert_eq!(read_dword_in(&store, REMOTE_IMAGES), Some(1));
        let text = std::fs::read_to_string(dir.0.join("mdview.ini")).unwrap();
        assert!(text.contains("NetworkFiles=2") && text.contains("RemoteImages=1"), "{text:?}");
    }

    // #16: the zoom is saved in this program's store; a value that isn't one
    // of the steps (edited by hand) gives the default.
    #[test]
    fn test_zoom_is_stored_and_unknown_values_give_100() {
        let dir = TempFolder::new("zoom");
        let store = ini_store_for(&dir.0);
        assert_eq!(zoom_in(&store), 100, "nothing saved yet");
        let failed = AtomicBool::new(false);
        assert_eq!(save_zoom_in(&store, 125, &failed), None);
        assert_eq!(zoom_in(&store), 125);
        let text = std::fs::read_to_string(dir.0.join("mdview.ini")).unwrap();
        assert!(text.contains("Zoom=125"), "{text:?}");
        for bad in [0, 105, 1000] {
            write_dword_in(&store, ZOOM, bad).unwrap();
            assert_eq!(zoom_in(&store), 100, "{bad}");
        }
    }

    // #50: fonts are stored per program; text values round-trip, an empty
    // name means MDView's default.
    #[test]
    fn test_font_settings_round_trip() {
        let dir = TempFolder::new("fonts");
        let store = ini_store_for(&dir.0);
        assert_eq!(font_settings_in(&store), (FontSettings::default(), vec![]), "nothing saved yet");
        let fonts = FontSettings {
            text_font: "Georgia".into(),
            text_size: 17,
            heading_font: String::new(),
            code_font: "Courier New".into(),
            code_size: 13,
        };
        assert!(set_font_settings_in(&store, &fonts).is_empty());
        assert_eq!(font_settings_in(&store), (fonts, vec![]));
        let text = std::fs::read_to_string(dir.0.join("mdview.ini")).unwrap();
        assert!(text.contains("TextFont=Georgia") && text.contains("TextSize=17") && text.contains("HeadingFont="), "{text:?}");
    }

    // A value edited by hand that could break out of the page's CSS, or a
    // size out of range, is not used, and the user is told.
    #[test]
    fn test_invalid_saved_fonts_are_not_used_and_noted() {
        let dir = TempFolder::new("badfonts");
        let path = dir.0.join("mdview.ini");
        std::fs::write(&path, "[MDView]\r\nTextFont=x\"; } body { display:none\r\nCodeFont=Consolas\r\nTextSize=400\r\nCodeSize=0\r\n").unwrap();
        let (fonts, notes) = font_settings_in(&Store::Ini(path));
        assert_eq!(fonts.text_font, "");
        assert_eq!(fonts.code_font, "Consolas");
        assert_eq!(fonts.text_size, DEFAULT_TEXT_SIZE);
        assert_eq!(fonts.code_size, 0);
        assert_eq!(notes.len(), 2, "{notes:?}");
    }

    #[test]
    fn test_font_names_and_sizes() {
        for ok in ["Segoe UI", "Cascadia Code", "Times New Roman", "MS Gothic", "Bahnschrift SemiBold", "Noto Sans CJK JP", "メイリオ", "Courier 10 Pitch", "A&B (x)"] {
            assert!(is_font_name(ok), "{ok}");
        }
        for bad in ["", " Arial", "Arial ", "a\"b", "a;b", "a{b", "a<b", "a\\b", "a/b", "a\nb", &"x".repeat(65)] {
            assert!(!is_font_name(bad), "{bad:?}");
        }
        assert!(is_font_size(8) && is_font_size(48) && !is_font_size(7) && !is_font_size(49));
    }

    // Saving names each value that could not be saved.
    #[test]
    fn test_font_save_names_what_failed() {
        let store = Store::Unavailable("no settings folder".to_string());
        let failed = set_font_settings_in(&store, &FontSettings::default());
        let labels: Vec<&str> = failed.iter().map(|(l, _)| *l).collect();
        assert_eq!(labels, ["Text font", "Text size", "Heading font", "Code font", "Code size"]);
        assert_eq!(labels.len(), FONT_VALUES);
        assert!(failed.iter().all(|(_, e)| e == "no settings folder"));
    }

    // A failed save is reported once per run, not for every zoom step.
    #[test]
    fn test_zoom_save_error_is_reported_once() {
        let store = Store::Unavailable("no settings folder".to_string());
        let failed = AtomicBool::new(false);
        assert_eq!(save_zoom_in(&store, 110, &failed).as_deref(), Some("no settings folder"));
        assert_eq!(save_zoom_in(&store, 125, &failed), None);
        assert_eq!(zoom_in(&store), 100);
    }
}
