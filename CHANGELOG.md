# Changelog

Each release's notes are taken from its section below (`## vX.Y.Z`). The
release workflow refuses to publish a version without a section here.

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
