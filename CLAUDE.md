# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

Personal academic homepage for Prof. Kichan PARK (박기찬), Professor Emeritus of Inha University. The entire site is a **single self-contained HTML file**: `park_kichan_homepage.html` (728 lines) with no build system, no dependencies, and no external tooling. To view the site, open the file directly in a browser.

## Architecture

The file has three co-located sections:

1. **CSS** (lines 8–182) — all styles are embedded in `<style>` using CSS custom properties defined in `:root`. The color palette is: `--navy` / `--navy-mid` / `--navy-light` / `--gold` / `--gold-light` / `--gold-pale` / `--cream`. Stick to these variables for any new styling.

2. **HTML** (lines 183–631) — two top-level divs gate the UX:
   - `#loginGate` — shown on load; hides the main content until the user authenticates.
   - `#mainSite` — the full profile page, revealed after login.

   Content sections inside `#mainSite` (in order): `#profile` (hero), `#philosophy`, `#education`, `#career`, `#research`, `#publications`, `#books`, `#projects`, plus a footer and an admin edit modal (`#editModal`).

3. **JavaScript** (lines 632–727) — vanilla JS, no frameworks. All functions are global. Key areas:
   - **Auth**: users stored in `localStorage` key `pkU` (array of `{n, e, o, p, a}` objects). Session stored in `sessionStorage` key `pkS`. An admin seed (`a: true`) for `kichan@inha.ac.kr` is auto-inserted on first load.
   - **Language toggle**: `data-lang` attribute on `<html>` switches between `ko` / `en`. Elements carry `data-ko` or `data-en` attributes, controlled by CSS rules at lines 184–188. Preference persisted in `localStorage` key `pkLang`.
   - **Admin edit**: The floating action button (`.edit-fab`) is visible only for admin users (`u.a === true`). It opens `#editModal`, which lets the admin edit the philosophy section text. Edits are persisted in `localStorage` keys `pkPhil_ko` and `pkPhil_en`.
   - **Publication filter**: `fp(cat, btn)` shows/hides `.pi2` items based on `data-cat` attribute.
   - **Scroll animations**: `IntersectionObserver` adds class `.v` to `.fi` elements when they enter the viewport.

## Bilingual Content Pattern

All user-visible text that needs translation uses paired elements:

```html
<span data-ko>한국어 텍스트</span>
<span data-en style="display:none">English text</span>
```

CSS at lines 184–188 handles visibility based on `html[data-lang]`. Always add both `data-ko` and `data-en` siblings when adding new visible text.

## Content Updates

All content (publications, career entries, education, projects) is hardcoded in HTML. There is no CMS or data file. To update a publication or add a career entry, edit the HTML directly in the relevant `<section>` block. Publication items use class `.pi2` with a `data-cat` attribute for filtering (e.g., `data-cat="ssci international"`).
