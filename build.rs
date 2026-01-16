fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").unwrap_or_default() == "windows" {
        let mut res = winresource::WindowsResource::new();

        // Version from Cargo.toml
        let version = std::env::var("CARGO_PKG_VERSION").unwrap_or_else(|_| "0.1.0".to_string());
        let version_parts: Vec<&str> = version.split('.').collect();
        let major = version_parts.first().unwrap_or(&"0");
        let minor = version_parts.get(1).unwrap_or(&"0");
        let patch = version_parts.get(2).unwrap_or(&"0");

        res.set("FileDescription", "MDView - Markdown Viewer")
            .set("ProductName", "MDView")
            .set("CompanyName", "Remko Weijnen")
            .set("LegalCopyright", "Copyright 2026 Remko Weijnen - Mozilla Public License 2.0")
            .set("OriginalFilename", "mdview.exe")
            .set("FileVersion", &format!("{}.{}.{}.0", major, minor, patch))
            .set("ProductVersion", &format!("{}.{}.{}.0", major, minor, patch));

        // Add icon if it exists
        if std::path::Path::new("assets/mdview.ico").exists() {
            res.set_icon("assets/mdview.ico");
        }

        if let Err(e) = res.compile() {
            eprintln!("cargo:warning=Failed to compile Windows resource: {}", e);
        }
    }
}
