//! Settings shared by the standalone viewer and the Total Commander plugin,
//! stored under `HKCU\Software\MDView`.

use windows::Win32::System::Registry::{
    HKEY, HKEY_CURRENT_USER, KEY_READ, KEY_WRITE, REG_DWORD, REG_OPTION_NON_VOLATILE,
    REG_VALUE_TYPE, RegCloseKey, RegCreateKeyExW, RegOpenKeyExW, RegQueryValueExW,
    RegSetValueExW,
};
use windows::core::PCWSTR;

pub const REGISTRY_KEY: &str = "Software\\MDView";

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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_remote_images_setting_round_trips_and_defaults_to_ask() {
        for s in [RemoteImagesSetting::Ask, RemoteImagesSetting::Always, RemoteImagesSetting::Never] {
            assert_eq!(RemoteImagesSetting::from_dword(s.to_dword()), s);
        }
        assert_eq!(RemoteImagesSetting::from_dword(99), RemoteImagesSetting::Ask);
    }
}
