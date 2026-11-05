# Changelog

Each release's notes are taken from its section below (`## vX.Y.Z`). The
release workflow refuses to publish a version without a section here.

## v0.2.7

### Fixes
- The standalone viewer and the Total Commander plugin can be used at the same time. Before, with MDView open in TC's Lister, a double-clicked `.md` showed nothing (and the other way round). WebView2 data now lives in `%LOCALAPPDATA%\MDView\WebView2`, one folder per program. (#33)
- This also fixes keyboard shortcuts such as **Ctrl+C** and **Esc** not working in the standalone viewer while MDView was open in TC. (#22, #23)
- **File → Register as .md Viewer** opens Windows Settings on MDView's default-apps page (Windows 11), instead of only showing "To change your default apps, go to Settings". If registering fails, MDView now says so. (#36)

### Changes
- The Total Commander plugin keeps its settings in its own `mdview.ini` in TC's settings folder (next to `wincmd.ini`), separate from the standalone viewer, which keeps using the registry. Answers given in TC (for example "Always load" remote images) apply to TC only. A choice made in the v0.2.6 plugin is asked once more. (#34)

## v0.2.6

### Hardening
MDView now treats Markdown files as untrusted content. Updating is recommended.
- **Scripts in documents don't run.** Markdown can contain HTML, including scripts. MDView only runs its own scripts and only shows its own pages, also when viewing with F3 in Total Commander.
- **Links don't start programs.** Links to programs and scripts (`.exe`, `.cmd`, `.ps1`, `.lnk`, …) are refused with a message. Other local files open after you confirm. Web and mail links open in your browser or mail app, as before.
- **Remote images are blocked by default** for privacy. A bar at the top offers **Load images**, **Always load** or **Never**.
- **Terminal output:** escape sequences in a document are not passed to the terminal.

### Other
- The README shows all views in one overview image.

## v0.2.5

### Fixes
- In-document links, such as a table of contents, now jump to the right heading. Headings get GitHub-compatible anchors, including in non-Latin scripts. (#4)
- YAML (`---`) and TOML (`+++`) front matter is no longer shown as part of the document. (#9, thanks @PurpleMyst)
- Standalone viewer: keyboard shortcuts (Ctrl+F, scrolling keys, Esc) work straight away, without clicking in the document first. (#18)
- Standalone viewer: no longer reopens minimised after being closed while minimised. A window that was maximised before minimising reopens maximised. (#6)

### New
- Standalone viewer: **Edit → Find… (Ctrl+F)**.

### Other
- The README no longer claims syntax highlighting. Code blocks are shown as monospaced text.
