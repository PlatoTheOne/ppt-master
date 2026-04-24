# Local Web UI

This repository now includes a small local web interface for users who prefer:

- upload a file
- paste content
- generate a draft first
- edit titles and slide order
- download `.pptx`
- export a single-file web presentation
- preview slides in the browser before opening PowerPoint
- reopen recent jobs from a history list
- open the project folder or exports folder in Explorer

## Start

From the repo root:

```powershell
python run_webui.py
```

Then open:

```text
http://127.0.0.1:8765
```

## What It Does

The Web UI wraps the existing local pipeline:

1. Create a project under `projects/`
2. Import or normalize source content into `sources/`
3. Build a draft slide outline
4. Render deterministic SVG draft slides into `svg_output/`
5. Run `svg_quality_checker.py`
6. Let the user edit deck title, subtitle, and slide order
7. Optionally run `finalize_svg.py`
8. Export PPTX via `svg_to_pptx.py`, Web Deck HTML, or both

Outputs are saved inside the generated project folder:

- native editable PPTX
- SVG reference PPTX
- single-file Web Deck HTML

## Input Modes

- Drag-and-drop upload: `.pdf`, `.docx`, `.md`, `.txt`, `.pptx`, `.html`, `.epub`, and other repo-supported formats
- Click-to-select upload
- Paste text directly
- Provide a URL

## Built-in Templates

The local UI currently exposes four visual directions:

- `business`
- `tech`
- `academic`
- `guizang`

These are still local deterministic renderers, but each one uses different theme tokens and slide framing so the decks are easier to distinguish at a glance.

The `guizang` preset borrows from `guizang-ppt-skill` in a pipeline-safe way:

- editorial pacing instead of repeating the same slide card pattern
- stronger chapter / quote / closing page rhythm
- magazine-like typography and margin-note framing
- still rendered to SVG locally and exported through the existing editable PPTX pipeline

## Browser Preview

After generation, the UI shows:

- a large slide viewer
- previous / next navigation
- thumbnail strip for direct page switching
- draft slides before export
- download links for both PPTX outputs after export

## Draft-first Workflow

The UI works in two phases:

1. Generate draft
2. Edit titles / reorder slides
3. Export PPTX, Web Deck, or both

It also includes:

- progress bar with stage label
- re-render draft action
- retry export action

## Export Modes

You can choose one of three export targets before generation or before export:

- `pptx`: keep the existing native PowerPoint export flow
- `web`: generate a browser-first single HTML deck
- `both`: write both outputs from the same edited draft

`Web Deck` is generated from the saved draft plan and theme. It does not replace the PPTX pipeline; it sits next to it as a second output mode.

## Web Deck Presentation Mode

The single-file `Web Deck` now includes a lightweight presentation layer:

- fullscreen toggle
- presenter mode with a side notes panel
- presenter console popup for a true two-screen setup
- presenter clock and talk timer
- next-slide preview with title, type, and key points
- keyboard shortcuts for navigation and stage controls
- top progress bar and current-slide status
- hide-UI mode for cleaner live presentation

Shortcuts inside the exported Web Deck:

- `Left` / `Right`: previous or next slide
- `F`: fullscreen
- `P`: presenter mode
- `O`: open the presenter console window
- `N`: notes panel
- `H`: hide or show presentation chrome
- `T`: start or pause presenter timer
- `R`: reset presenter timer

For a two-screen talk flow:

1. Open the exported `Web Deck`
2. Press `O` or click `Console`
3. Move the popup to your presenter display
4. Keep the main deck on the audience display

## History And Folder Actions

The UI also shows a recent history list built from local `projects/` output folders.

From the interface you can:

- reopen a previous deck in the preview area
- download its latest PPTX exports
- open the project directory in Explorer
- open the exports directory in Explorer

## Planner Modes

### Heuristic

- No API key required
- Uses simple local structure extraction
- Best for quick smoke tests and structured source material

### OpenAI-compatible

- Requires API key
- Optional custom base URL
- Uses a model only for slide planning
- Final SVG rendering and PPT export still happen locally

If the OpenAI planner fails, the UI falls back to heuristic planning so the job can still finish.

## Current Scope

This UI is intentionally minimal:

- focused on local generation
- no multi-user auth
- no persistent database
- no cloud storage
- no image-generation workflow yet

It is a practical wrapper around the repo, not a full SaaS frontend.
