# MDView

A fast, lightweight Markdown viewer for Windows. Available as a Total Commander Lister plugin and standalone executable.

![MDView Screenshot](assets/screenshot.png)

## Features

- **WebView2 rendering** - Modern HTML rendering with full Markdown support
- **Ctrl+click links** - Open external links in your default browser
- **ESC to close** - Quick keyboard navigation
- **Syntax highlighting** - Code blocks with proper formatting
- **GitHub Flavored Markdown** - Tables, task lists, strikethrough, and more

## Installation

### Total Commander Plugin

1. Download the latest release ZIP
2. Open the ZIP file in Total Commander
3. Total Commander will automatically detect `pluginst.inf` and offer to install
4. Select your preferred installation directory
5. Configure file associations (`.md`, `.markdown`) in TC settings

### Standalone Executable

1. Download `mdview.exe` from the latest release
2. Place it anywhere in your PATH or desired location
3. Associate `.md` files with `mdview.exe` or run from command line

## Usage

### Command Line

```
mdview [OPTIONS] [FILE]

Arguments:
  [FILE]  Markdown file to view (reads from stdin if not provided)

Options:
  -h, --help     Print help
  -V, --version  Print version
```

### Examples

```bash
# View a markdown file
mdview README.md

# Pipe content from another command
cat notes.md | mdview

# View with explicit file
mdview --file documentation.md
```

### Keyboard Shortcuts

| Key | Action |
|-----|--------|
| ESC | Close viewer |
| Ctrl+Click | Open link in browser |

## Building from Source

### Prerequisites

- Rust 1.75 or later
- Windows 10/11 with WebView2 Runtime

### Build Commands

```bash
# Build release version (x64)
cargo build --release

# Build x86 version (for 32-bit Total Commander)
cargo build --release --target i686-pc-windows-msvc

# Run tests
cargo test
```

### Output Files

After building, copy the following files for distribution:

| File | Description |
|------|-------------|
| `target/release/mdview.exe` | Standalone viewer (x64) |
| `target/release/mdview_wlx.dll` | Rename to `mdview.wlx64` for TC plugin (x64) |
| `target/i686-pc-windows-msvc/release/mdview_wlx.dll` | Rename to `mdview.wlx` for TC plugin (x86) |

## Requirements

- Windows 10 version 1803 or later
- [WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/) (usually pre-installed on Windows 10/11)

## License

This project is licensed under the Mozilla Public License 2.0 - see the [LICENSE](LICENSE) file for details.

## Author

Remko Weijnen

## Acknowledgments

- [pulldown-cmark](https://github.com/raphlinus/pulldown-cmark) - Markdown parsing
- [webview2-com](https://github.com/nicksenger/webview2-com) - WebView2 bindings for Rust
