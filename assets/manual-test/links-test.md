# MDView Link Test

Use this file to manually test link handling in both:
- `mdview.exe`
- Total Commander lister plugin

## External links

- [Rust website](https://www.rust-lang.org/)
- [GitHub](https://github.com/)
- [Send mail](mailto:test@example.com)

## External image link

Click the badge itself:

[![Rust badge](https://img.shields.io/badge/Rust-stable-orange)](https://www.rust-lang.org/)

## Relative markdown links

- [Open second.md](second.md)
- [Open nested/third.md](nested/third.md)
- [Open file with spaces](file%20with%20spaces.md)
- [Open README.md from repo root](../../README.md)

## Relative non-markdown files

- [Open LICENSE](../../LICENSE)
- [Open screenshot image](../screenshot.png)

## Fragment links

- [Jump to Section A](#section-a)
- [Jump to Section B](#section-b)
- [Jump to STEP 1 — Architecture Analysis](#step-1--architecture-analysis) (GitHub-style anchor, issue #4)
- [Jump to Βήμα 2 — Ανάλυση](#βήμα-2--ανάλυση) (non-ASCII heading)

## Mixed links with fragments

- [Open second.md#details](second.md#details)
- [Open nested/third.md#bottom](nested/third.md#bottom)

## Inline image

If local resource mapping works, this image should render:

![Local screenshot](../screenshot.png)

## Section A

Some text here.

## Section B

Some more text here.

## STEP 1 — Architecture Analysis

Target of the GitHub-style anchor link.

## Βήμα 2 — Ανάλυση

Target of the non-ASCII anchor link.
