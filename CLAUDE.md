# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

Company presentation assets for **Oceans5 Group AS** — a Norwegian marine engineering and AI project management consultancy. The repo contains:

- `oceans5-deck/` — the primary deliverable: a 12-slide reveal.js HTML presentation
- `generate_pptx.py` — generates a matching PPTX from scratch using python-pptx
- `research/` — text extracted from source documents (strategy doc, existing PPTX)
- Source files (gitignored): `*.docx`, `*.pptx`, `*.potx`, `*.thmx`, `*.jpg`

## Running the presentation

Open `oceans5-deck/index.html` directly in any modern browser — no build step, no server needed. All dependencies load from CDN (reveal.js 5.1, Chart.js 4.4, Google Fonts).

## Generating the PPTX

```bash
pip install python-pptx   # once
python generate_pptx.py
```

Outputs `Oceans5_Company_Presentation_2026_02_English.pptx` in the project root. The file is gitignored (binary).

## Deploying to GitHub Pages

```bash
gh repo create oceans5-company-presentation --public --source . --remote origin --push
# Then enable Pages: Settings → Pages → Source: master / root
```

The live deck URL will be: `https://{username}.github.io/oceans5-company-presentation/oceans5-deck/`

## Architecture

### HTML deck (`oceans5-deck/`)

- `index.html` — all 12 slides inline. Each slide is a `<section>` in reveal.js. Slides 5, 9, 10 contain Chart.js canvases initialised on `slidechanged` and `ready` events.
- `css/oceans5-theme.css` — all brand styling. CSS custom properties define the palette at `:root`. Layout uses CSS Grid (`grid-2`, `grid-3`, `grid-4`) and flexbox. No external CSS dependencies beyond Google Fonts.

**Slide index (0-based for JS):** Cover=0, Challenge=1, Who We Are=2, Services=3, AI Advantage=4, Track Record=5, Approach=6, Team=7, Market=8, Vision 2028=9, Why Oceans5=10, CTA=11.

**Charts** live in slides 4, 8, 9 — each has a `init*Chart()` function gated by a boolean flag to prevent double-init on back-navigation.

### PPTX generator (`generate_pptx.py`)

One function per slide (`slide_01_cover`, `slide_02_challenge`, …). All helpers (`add_tb`, `add_rect`, `add_card`, `add_tag`, `add_footer`, `accent_bar`) are at the top. Brand colours are module-level `RGBColor` constants. Charts use `pptx.chart.data.ChartData` + `XL_CHART_TYPE`. Slide canvas is 13.333 × 7.5 inches (16:9).

## Brand

| Token | Hex |
|-------|-----|
| `--navy-deep` | `#0A1B2E` |
| `--teal` | `#007EA7` |
| `--teal-bright` | `#00B4D8` |
| `--off-white` | `#F4F8FB` |
| `--gold` | `#F0A500` |

Primary fonts: **Montserrat** (headings, bold), **Inter** (body). Both loaded from Google Fonts in the HTML; match these in any PPTX additions.

## Content source

Real company data lives in `research/strategy_extract.txt` (from `Konsernstrategi.docx`) and `research/pptx_extract.txt` (from the 2026 English PPTX). When updating slide copy, draw from these files — do not invent project names, metrics, or personnel.

Key facts to keep accurate:
- Founded 2021, Oslo · 3 companies · 5 co-owner leaders
- CEO: Peter-Emil Johannessen (`post@oceans5.no`, +47 412 69 164)
- Vision 2028 target: 30–40% revenue from work packages/services, ISO 9001, AI in all workflows
- Named projects: Empire Wind, Estlink, Dogger Bank, Baltyk 2&3, Tyrrhenian Link, Borwin5, Wilhelmshaven LNG Terminal, and others in `pptx_extract.txt`
