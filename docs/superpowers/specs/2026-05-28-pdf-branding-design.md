# PDF Branding — Design Spec

**Date:** 2026-05-28
**Status:** Approved

## Overview

Add the DS logo and preferred brand fonts (Merriweather + Roboto) to both PDF templates. Both the Plans and Pipeline PDFs currently use Arial throughout and have a text-only blue header. This change replaces that header with a split logo/date header and applies the brand font pair consistently.

---

## Files changed

| File | Change |
|---|---|
| `reports/plans_pdf.py` | `HTML(string=html)` → `HTML(string=html, base_url=str(_TEMPLATES_DIR))` |
| `reports/pipeline_pdf.py` | Same one-line change |
| `templates/plans_pdf.html` | New header HTML + `@font-face` declarations + font assignments |
| `templates/pipeline_pdf.html` | Same |

## New files added

`templates/fonts/` — five TTF files downloaded from Google Fonts:

- `Roboto-Regular.ttf`
- `Roboto-SemiBold.ttf`
- `Roboto-Bold.ttf`
- `Merriweather-Regular.ttf`
- `Merriweather-Bold.ttf`

Fonts live under `templates/` so all relative paths resolve from the same `base_url` as the logo image. WeasyPrint resolves relative URLs in `HTML(string=…)` mode only when a `base_url` is supplied; without it, relative `src` and `@font-face` `src` values silently fail.

---

## Header redesign

Both templates share the same header pattern. The existing single-colour blue banner is replaced with a two-cell split:

| Cell | Background | Content |
|---|---|---|
| Left | White `#ffffff` | `DS_Logo_Alt.png` at `40px` height |
| Right | Blue `#4472C4` | "Generated {date}" — Roboto, 9pt, light blue `#c5d5ee` |

The logo is referenced as `./DS_Logo_Alt.png` (relative to `base_url = _TEMPLATES_DIR`). The existing text-only branding line ("Diana Silver Grants Strategy · diana@dsilvergrants.com") is removed from the header — the logo carries the brand identity.

The `@page` footer rules (`@bottom-left`, `@bottom-right`) are unchanged in content but update `font-family` from `Arial, sans-serif` to `Roboto, sans-serif`.

---

## Font assignments

### `@font-face` declarations (identical in both templates)

```css
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Roboto; src: url('./fonts/Roboto-Bold.ttf'); font-weight: 700; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Regular.ttf'); font-weight: 400; }
@font-face { font-family: Merriweather; src: url('./fonts/Merriweather-Bold.ttf'); font-weight: 700; }
```

Both templates set `body { font-family: Roboto, sans-serif; }` as the base so all elements inherit Roboto by default. Merriweather is applied only to the specific heading selectors listed below.

### Plans PDF

| Element | Font | Weight |
|---|---|---|
| Client name (`.client-name`) | Merriweather | 700 |
| "Plans Report" label (`.report-type`) | Merriweather | 700 |
| Year section labels (`.year-label`) | Merriweather | 700 |
| Status group labels (`.status-label`) | Roboto | 700 |
| Grant row title (`.grant-title`) | Roboto | 700 |
| Grant row amounts (`.grant-amounts`) | Roboto | 400 |
| Grant detail line (`.grant-details`) | Roboto | 400 |
| Header date | Roboto | 400 |
| `@page` footer | Roboto | 400 |

### Pipeline PDF

| Element | Font | Weight |
|---|---|---|
| Client name (`.client-name`) | Merriweather | 700 |
| "Pipeline Report" label (`.report-type`) | Merriweather | 700 |
| Table header row (`th`) | Roboto | 700 |
| Table data cells (`td`) | Roboto | 400 |
| Header date | Roboto | 400 |
| `@page` footer | Roboto | 400 |

---

## No other changes

- **Dockerfile** — no changes. Fonts are committed files, not apt packages.
- **Tests** — no changes. Existing tests assert `%PDF-` magic bytes; the `base_url` addition is additive and does not affect test fixtures or assertions.
- **`app.py`** — no changes.
- **`templates/index.html`** — no changes.
