# Gabriel Elorde Portfolio

[![Cross-platform smoke test](https://github.com/maidenator/Streamlit-Portfolio/actions/workflows/smoke.yml/badge.svg)](https://github.com/maidenator/Streamlit-Portfolio/actions/workflows/smoke.yml)

A two-page portfolio for **Gabriel Elorde** ([@maidenator](https://github.com/maidenator)),
built with Streamlit and custom HTML/CSS as a Microsoft Word 2003 document
running on a Windows XP desktop.

The whole site lives inside a maximized Word window: title bar, menus,
toolbars, ruler, document map, status bar, and an XP taskbar. The document
contains a profile, the featured Artifact Speculation Engine, selected
repositories, education, toolkit, and contact details.

## Run

Requires Python 3.10 or newer.

**macOS and Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

**Windows PowerShell**

```powershell
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

**Windows Command Prompt**

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

If PowerShell blocks the activation script, run
`Set-ExecutionPolicy -Scope Process Bypass` in that window first. On a
headless server, add `--server.headless true` and open port 8501 from your
browser.

The app opens at <http://localhost:8501>. `runOnSave = true` is set in
`.streamlit/config.toml`, so the browser reloads when the page changes.

All file paths are resolved relative to `app.py`, and the static assets use
portable URLs. The GitHub Actions workflow in `.github/workflows/smoke.yml`
starts the app on Ubuntu, Windows, and macOS and checks the health endpoint
and every static asset.

Static file serving is enabled in `.streamlit/config.toml`
(`enableStaticServing = true`). The interactive Word layer is loaded from
`static/word-ui.js`; keep that file beside `app.py` when deploying.

The document text is set in **Times New Roman** and the interface chrome in
**Verdana** (with Tahoma, Segoe UI, and Arial fallbacks). Both fonts ship with
Windows and macOS; on Linux, install the Microsoft core fonts for the intended
look.

## What works

**Document**

- Click anywhere in the page to edit. Typing marks the title bar with `*`.
- Select text, then use **Bold**, *Italic*, underline, alignment, lists,
  indenting, highlight, font color, or the font/size boxes.
- **Format → Font**, **Format → Style**, **Format → Borders and Shading**,
  and the font-color picker open Word-style dialogs.
- **Insert** adds hyperlinks, dates, horizontal rules, page breaks, and
  tables. **Table** can insert/delete rows and columns or select a table.
- **Tools → Find**, **Word Count**, **Spelling and Grammar**, **Options**,
  **Language**, and **Portfolio Properties** are interactive.
- The **Document Map** links jump to sections; its filter switches between
  headings and pages. The ruler, toolbars, and Document Map can be toggled
  from **View**.
- **File → Save** downloads a Word-compatible HTML document with a `.doc`
  extension (styled for letter-size pages). **Save As** lets you rename it.
- **File → Open** opens plain text or HTML files as a simplified document.
  Binary `.doc` files are not parsed in the browser. **File → Revert to
  Portfolio** restores the original document.
- **File → Print** and **Ctrl+P** print only the two document pages.

**Window and desktop**

- Minimize, restore-down, maximize, and close the Word window from the title
  bar, the **Window** menu, or the taskbar. Closed documents can be reopened
  from the desktop icon.
- When the document is closed, drag the `Portfolio.doc` icon anywhere on the
  desktop with a mouse or touch. A plain click still reopens it.
- The desktop wallpaper is served from `static/bliss.jpg` and applied behind
  the XP desktop.
- The **Start** menu opens documents, GitHub, help, the Document Map, and a
  simulated shutdown/restart screen.
- Tray icons and the taskbar clock show XP-style balloon notifications.
- The collapsible paperclip assistant links to the help dialog.
- Keyboard shortcuts include `Ctrl+S`, `Ctrl+O`, `Ctrl+N`, `Ctrl+P`,
  `Ctrl+F`, `Ctrl+K`, `Ctrl+B/I/U`, `Ctrl+Z/Y`, `Ctrl+Enter`, `F1`, `F7`,
  `F11`, `F12`, and `Alt+F` for the File menu.

## Responsive and print

Below 900 px the menu bar, toolbars, ruler, and document map are replaced by a
compact section nav. Dialogs, the Start menu, and the taskbar stay usable down
to 320 px, and the page has no horizontal overflow.

The print stylesheet hides all Windows/Word chrome and prints the two
letter-size document pages. A saved `.doc` uses the same page breaks.

## Structure

```text
app.py                  # content, project data, and generated icon CSS
static/word-ui.js       # menus, dialogs, editing, save/open, window actions
static/bliss.jpg        # Windows XP desktop wallpaper
static/clippy.png       # paperclip assistant artwork
styles.css              # Word 2003 / XP chrome, document, responsive, print
scripts/smoke_test.py   # cross-platform startup and asset check
.github/workflows/smoke.yml  # Ubuntu, Windows, macOS smoke test
requirements.txt        # dependencies
.streamlit/config.toml  # Streamlit shell theme and server settings
```

## Customize

| What | Where |
| --- | --- |
| Name, intro copy, section content | `app.py` (`build_page`) |
| Project rows | `PROJECTS` in `app.py` |
| Social, featured, and GitHub stats | Constants at the top of `app.py` |
| Menu definitions and toolbar markup | `MENUS`, `tb_button`, `fmt_button`, `combo_button` in `app.py` |
| Toolbar/icon artwork | `ICONS`, `WORD_ICON_SVG`, `START_ICON_SVG` in `app.py` |
| Paperclip assistant | `static/clippy.png` |
| Interactive behavior | `static/word-ui.js` |
| Desktop wallpaper | `static/bliss.jpg` and `.xp-desktop` in `styles.css` |
| Layout, colors, type, responsive and print rules | `styles.css` |
| Streamlit shell colors | `.streamlit/config.toml` |

Toolbar, window, and tray artwork is inline SVG in `app.py`, converted to CSS
data URIs at render time. Streamlit's `st.html` strips literal `<svg>` tags, so
do not place raw SVGs in the page markup. Add them to the icon dictionary and
let `build_icon_css()` handle it. The paperclip assistant is the PNG at
`static/clippy.png`.

`st.html` also removes `contenteditable`, which is why `word-ui.js` applies it
to the document at startup. The interactive layer is served as a static file
rather than an inline script because the HTML sanitizer drops inline scripts
whose source contains HTML markup.

Project data is kept in `app.py` rather than fetched at runtime, so the page
stays fast and works without hitting the GitHub API.

## Deploy

Push the folder to GitHub, then create an app on
[Streamlit Community Cloud](https://share.streamlit.io) with `app.py` as the
entry point. Keep `static/word-ui.js`, `static/bliss.jpg`, `static/clippy.png`,
and `enableStaticServing = true`; no build step or extra assets are required.
