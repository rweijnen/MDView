//! Settings of the standalone viewer and the Total Commander plugin.
//!
//! Each program keeps its own settings, and answers to bars and prompts are
//! stored where they were given (#34):
//! - standalone viewer: `HKCU\Software\MDView`
//! - TC plugin: `mdview.ini` in TC's ini folder (next to wincmd.ini), as
//!   reported by TC through `ListSetDefaultParams` (WLX SDK)

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, KEY_READ, KEY_WRITE, REG_DWORD, REG_OPTION_NON_VOLATILE,
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

/// Whether remote (http/https) images in a document are loaded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RemoteImagesSetting {
    /// Default: blocked, with a banner offering to load them.
    Ask,
    /// Always load remote images.
    Always,
    /// Never load them and don't ask.
    Never,
}

impl RemoteImagesSetting {
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

pub fn remote_images_setting() -> RemoteImagesSetting {
    read_dword(REMOTE_IMAGES)
        .map(RemoteImagesSetting::from_dword)
        .unwrap_or(RemoteImagesSetting::Ask)
}

pub fn set_remote_images_setting(setting: RemoteImagesSetting) -> Result<(), String> {
    write_dword(REMOTE_IMAGES, setting.to_dword())
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
        for s in [RemoteImagesSetting::Ask, RemoteImagesSetting::Always, RemoteImagesSetting::Never] {
            assert_eq!(RemoteImagesSetting::from_dword(s.to_dword()), s);
        }
        assert_eq!(RemoteImagesSetting::from_dword(99), RemoteImagesSetting::Ask);
    }
}
