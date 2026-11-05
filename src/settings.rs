//! Settings shared by the standalone viewer and the Total Commander plugin,
//! stored under `HKCU\Software\MDView`.

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

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Read a DWORD value; `None` if the key or value does not exist.
pub fn read_dword(name: &str) -> Option<u32> {
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
pub fn write_dword(name: &str, value: u32) -> Result<(), String> {
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

    #[test]
    fn test_remote_images_setting_round_trips_and_defaults_to_ask() {
        for s in [RemoteImagesSetting::Ask, RemoteImagesSetting::Always, RemoteImagesSetting::Never] {
            assert_eq!(RemoteImagesSetting::from_dword(s.to_dword()), s);
        }
        assert_eq!(RemoteImagesSetting::from_dword(99), RemoteImagesSetting::Ask);
    }
}
