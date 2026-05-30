# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

Company presentation assets for **Oceans5 Group AS** — a Norwegian marine engineering and AI project management consultancy. The repo contains:

- `oceans5-deck/` — the primary deliverable: a 12-slide reveal.js HTML presentation
- `generate_pptx.py` — generates a matching PPTX from scratch using python-pptx
- `research/` — text extracted from source documents (strategy doc, existing PPTX)
- `o5-design-system/` — official brand design system (gitignored zip + extracted folder)
- Source files (gitignored): `*.docx`, `*.pptx`, `*.potx`, `*.thmx`

## Running the presentation

Open `oceans5-deck/index.html` directly in any modern browser — no build step, no server needed. All dependencies load from CDN (reveal.js 5.1, Chart.js 4.4, Google Fonts).

Live URL (GitHub Pages): `https://tdn-dotcom.github.io/oceans5-company-presentation/oceans5-deck/`

## Generating the PPTX

```bash
pip install python-pptx   # once
python generate_pptx.py
```

Outputs `Oceans5_Company_Presentation_2026_02_English.pptx` in the project root. The file is gitignored (binary).

## Architecture

### HTML deck (`oceans5-deck/`)

- `index.html` — all 12 slides inline. Each slide is a `<section>` in reveal.js.
- `css/oceans5-theme.css` — all brand styling. CSS custom properties at `:root` mirror the official design tokens from `colors_and_type.css`. Layout uses utility classes: `grid-2`, `grid-3`, `grid-4`, `grid-2-1`. Card variants: `card-light` (white, indigo shadow), `card-inv` (deep-indigo, cyan overline), `card-glass` (translucent, for use over photos).
- `assets/logo-wordmark-white.svg` / `logo-wordmark-indigo.svg` — official logo vectors. Never redraw or replace with text.
- `assets/photography/` — 5 approved brand photos used as full-bleed slide backgrounds.

**Slide index (0-based for JS):** Cover=0, Challenge=1, Who We Are=2, Services=3, AI Advantage=4, Track Record=5, Approach=6, Team=7, Market=8, Vision 2028=9, Why Oceans5=10, CTA=11.

**Photo backgrounds** use `.photo-bg` + a wash modifier (`.wash-bottom`, `.wash-full`, `.wash-dark`) to apply an indigo overlay per brand guidelines. Slides with photos: Cover (`offshore-wind-sunset.jpg`), AI Advantage (`night-cable-lay.jpg`), Track Record (`aerial-cable-lay.jpg`), Vision 2028 (`platform-dusk.jpg`), CTA (`crew-survey.jpg`).

**Charts** live in slides 4, 8, 9 — each has an `init*Chart()` function gated by a boolean flag (prevents double-init on back-navigation) and fires on both `slidechanged` and `ready` events (handles direct URL hash navigation).

### PPTX generator (`generate_pptx.py`)

One function per slide (`slide_01_cover` … `slide_12_cta`). Helpers at top: `add_tb`, `add_rect`, `add_card`, `add_tag`, `add_footer`, `accent_bar`. Brand colours are module-level `RGBColor` constants — update these to match the HTML tokens if the palette changes. Charts use `pptx.chart.data.ChartData` + `XL_CHART_TYPE`. Slide canvas is 13.333 × 7.5 inches (16:9).

## Brand (official design system)

Source of truth: `o5-design-system/design_handoff_oceans5_design_system/colors_and_type.css` and `README.md`.

### Colours
| Token | Hex | Role |
|-------|-----|------|
| `--deep-indigo` | `#2D3282` | Primary brand — backgrounds, headings |
| `--azure-blue` | `#3F75FF` | Digital accent — links, CTAs, chart lines |
| `--saffron-yellow` | `#FEB929` | Highlight — sparing use only |
| `--cyan` | `#2BF7FF` | Highlight on dark — overlines on inverse cards |
| `--eclipse` | `#0E0E0D` | Ink / near-black |
| `--alabaster` | `#FAF9F7` | Default warm surface (light slides) |
| `--indigo-80/50/30` | `#575b9b / #9698c0 / #c0c1da` | Tints — muted text, disabled |
| `--heritage-steel` | `#355283` | Original logo blue — heritage use |

Semantic: `--bg`=alabaster · `--bg-inverse`=deep-indigo · `--fg`=eclipse · `--fg-on-dark`=alabaster · `--accent`=azure · `--highlight`=saffron.

### Typography
- **Display/headings:** `Sora` (Google Fonts substitute for licensed brand font **Rustica**). Weight 500 for H1–H4. Letter-spacing `-0.01em`.
- **Body/UI:** `Hanken Grotesk` (substitute for **Greycliff CF**). Weight 400 body, 600 labels.
- **Technical:** `Space Mono`.
- When licensed Rustica + Greycliff CF webfonts are supplied, swap the `@import` in `oceans5-theme.css` and update `--font-display` / `--font-text`.

### Logo
Use only the SVGs in `oceans5-deck/assets/`. Do not recreate the wordmark in code or use plain text. Clearspace = one cap-height on all sides. White wordmark on dark/photo; indigo wordmark on light.

### Photography
Five approved photos in `oceans5-deck/assets/photography/`. On full-bleed hero images, reverse the white wordmark out of an indigo wash gradient at the base (`.photo-bg.wash-bottom`). Do not use stock images or apply warm filters.

## Content source

Real company data lives in `research/strategy_extract.txt` (from `Konsernstrategi.docx`) and `research/pptx_extract.txt` (from the 2026 English PPTX). When updating slide copy, draw from these files — do not invent project names, metrics, or personnel.

Key facts to keep accurate:
- Founded 2021, Oslo (Drammensveien 126A, 0277) · 3 companies · 5 co-owner leaders
- CEO: Peter-Emil Johannessen (`post@oceans5.no`, +47 412 69 164)
- Vision 2028 ("Nordstjernen"): 30–40% revenue from work packages/services, ISO 9001, AI in all workflows
- Named projects: Empire Wind, Estlink, Dogger Bank, Baltyk 2&3, Tyrrhenian Link, Borwin5, Wilhelmshaven LNG Terminal, and others in `pptx_extract.txt`
