from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import streamlit as st

st.set_page_config(
    page_title="Portfolio.doc - Microsoft Word",
    page_icon="📄",
    layout="wide",
)

BASE_DIR = Path(__file__).parent
GITHUB_URL = "https://github.com/maidenator"
REPOSITORIES_URL = f"{GITHUB_URL}?tab=repositories"
FEATURED_LIVE_URL = "https://artifact-speculation-engine.vercel.app"
FEATURED_REPO_URL = f"{GITHUB_URL}/GI-Artifact-Speculation-Engine"

PUBLIC_REPO_COUNT = 15

PROJECTS = [
    {
        "title": "Zen Chrome Spoofer",
        "description": (
            "A small browser extension for Zen and Firefox that spoofs the user "
            "agent, so websites think you're browsing in Chrome."
        ),
        "tags": ["JavaScript", "Browser extension"],
        "url": f"{GITHUB_URL}/zen-chrome-spoofer",
    },
    {
        "title": "DAA Problemsets",
        "description": (
            "A collection of algorithm problems and solutions from my Design "
            "and Analysis of Algorithms course."
        ),
        "tags": ["C++", "Algorithms"],
        "url": f"{GITHUB_URL}/DAA-Problemsets",
    },
    {
        "title": "Artifact Engine · Blazor Port",
        "description": (
            "The artifact simulator rebuilt in C# and Blazor."
        ),
        "tags": ["C#", "Blazor"],
        "url": f"{GITHUB_URL}/gi-artifact-speculation-engine-blazor-port",
    },
    {
        "title": "Tic-Tac-Toe Minimax",
        "description": "A tic-tac-toe game with a minimax opponent.",
        "tags": ["C#", "Game AI"],
        "url": f"{GITHUB_URL}/TicTacToeMinimax",
    },
    {
        "title": "Back-Propagation",
        "description": (
            "A from-scratch C# implementation of back-propagation for training "
            "a neural network."
        ),
        "tags": ["C#", "Neural networks"],
        "url": f"{GITHUB_URL}/Back-Propagation",
    },
]


def svg_16(inner: str) -> str:
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16">'
        f"{inner}</svg>"
    )


WORD_ICON_SVG = svg_16(
    '<path d="M2.5 1h7L13 4.5V15H2.5z" fill="#fff" stroke="#063d91"/>'
    '<path d="M9.5 1v3.5H13" fill="#dbe8ff" stroke="#063d91"/>'
    '<path d="M4.6 6.2 5.7 12h1.1L8 7.7 9.2 12h1.1l1.1-5.8h-1.2l-.6 4.1-1.1-4.1H7.4L6.3 10.3l-.6-4.1z" fill="#1857b8"/>'
)

START_ICON_SVG = svg_16(
    '<rect x="1" y="2" width="6" height="6" fill="#fff"/>'
    '<rect x="9" y="2" width="6" height="6" fill="#bfe0ad"/>'
    '<rect x="1" y="10" width="6" height="6" fill="#bfe0ad"/>'
    '<rect x="9" y="10" width="6" height="6" fill="#fff"/>'
)

ICONS = {
    "new": '<path d="M4 1.5h5.5L13 5v9.5H4z" fill="#fff" stroke="#444"/><path d="M9.5 1.5V5H13" fill="none" stroke="#444"/>',
    "open": '<path d="M1.5 4h5l1 1.5h7V13h-13z" fill="#f3d27a" stroke="#7a5b12"/><path d="M1.5 6.5h13" stroke="#7a5b12"/>',
    "save": '<path d="M2.5 2.5h9l2 2v9h-11z" fill="#e9e9e9" stroke="#444"/><path d="M5 2.5h6v3.5H5z" fill="#fff" stroke="#777"/><path d="M5 9.5h6V13H5z" fill="#2f6bb0"/>',
    "print": '<path d="M4 2.5h8v3H4z" fill="#fff" stroke="#444"/><path d="M2 5.5h12v5H2z" fill="#c9c9c9" stroke="#444"/><path d="M4.5 9h7v4.5h-7z" fill="#fff" stroke="#444"/><path d="M11.5 7h1" stroke="#2f8f3b" stroke-width="1.5"/>',
    "undo": '<path d="M3 7.5h6.5a3.5 3.5 0 0 1 0 7H6" fill="none" stroke="#2f5f9f" stroke-width="1.6"/><path d="M5.5 5 2.5 7.5 5.5 10" fill="none" stroke="#2f5f9f" stroke-width="1.6"/>',
    "redo": '<path d="M13 7.5H6.5a3.5 3.5 0 0 0 0 7H10" fill="none" stroke="#2f5f9f" stroke-width="1.6"/><path d="M10.5 5l3 2.5-3 2.5" fill="none" stroke="#2f5f9f" stroke-width="1.6"/>',
    "link": '<path d="M6.2 4.2H4.5a3.2 3.2 0 0 0 0 6.4h1.7M9.8 4.2h1.7a3.2 3.2 0 0 1 0 6.4H9.8M5 7.4h6" fill="none" stroke="#1f4f92" stroke-width="1.5"/><path d="M3.5 2.5h3v1.4h-3z" fill="#1f4f92"/>',
    "table": '<rect x="2" y="3" width="12" height="10" fill="#fff" stroke="#444"/><path d="M2 6.5h12M2 10h12M6 3v10M10 3v10" stroke="#777"/>',
    "spell": '<path d="M2 12.5 5 4h1.4l3 8.5H8.1l-.7-2.2H4.6l-.7 2.2z" fill="#1f4f92"/><path d="M5.1 8.8h2.3L6.2 6z" fill="#fff"/><path d="M10 10.5l1.3 1.4 2.6-3" fill="none" stroke="#2f8f3b" stroke-width="1.6"/>',
    "zoom": '<circle cx="7" cy="7" r="4.2" fill="#fff" stroke="#444"/><path d="M10.2 10.2 14 14" stroke="#444" stroke-width="2"/><path d="M5 7h4M7 5v4" stroke="#2f5f9f"/>',
    "align-left": '<path d="M2 3.5h12M2 6.5h8M2 9.5h12M2 12.5h8" stroke="#333" stroke-width="1.4"/>',
    "align-center": '<path d="M2 3.5h12M4 6.5h8M2 9.5h12M4 12.5h8" stroke="#333" stroke-width="1.4"/>',
    "align-right": '<path d="M2 3.5h12M6 6.5h8M2 9.5h12M6 12.5h8" stroke="#333" stroke-width="1.4"/>',
    "align-justify": '<path d="M2 3.5h12M2 6.5h12M2 9.5h12M2 12.5h12" stroke="#333" stroke-width="1.4"/>',
    "bullets": '<path d="M5.5 3.5h9M5.5 8h9M5.5 12.5h9" stroke="#333" stroke-width="1.3"/><path d="M2 3.5h1.5M2 8h1.5M2 12.5h1.5" stroke="#333" stroke-width="2.2"/>',
    "numbers": '<path d="M6 3.5h8M6 8h8M6 12.5h8" stroke="#333" stroke-width="1.3"/><text x="1.5" y="5" font-family="Tahoma" font-size="4.5" fill="#333">1</text><text x="1.5" y="9.5" font-family="Tahoma" font-size="4.5" fill="#333">2</text><text x="1.5" y="14" font-family="Tahoma" font-size="4.5" fill="#333">3</text>',
    "indent": '<path d="M6 3.5h8M6 8h8M6 12.5h8" stroke="#333" stroke-width="1.3"/><path d="M2 5.5 4.5 8 2 10.5z" fill="#2f5f9f"/><path d="M2 3.5v9" stroke="#333"/>',
    "border": '<rect x="2" y="2.5" width="12" height="11" fill="none" stroke="#555"/><path d="M4 4.5h8v7H4z" fill="none" stroke="#aaa" stroke-dasharray="1.5 1.5"/>',
    "highlight": '<path d="M10 2.5 13.5 6 7 12.5 3.5 9z" fill="#ffe36a" stroke="#7a6410"/><path d="M3.5 9 7 12.5 5 14.5H2v-3z" fill="#d9b62b" stroke="#7a6410"/><path d="M2 15h12" stroke="#7a6410" stroke-width="1.4"/>',
    "font-color": '<path d="M3 11.5 6 3h1.5l3 8.5H9l-.7-2.2H5.6L5 11.5z" fill="#333"/><path d="M6 7.8h2.3L7.1 5z" fill="#fff"/><path d="M2.5 14h11" stroke="#c22" stroke-width="2"/>',
    "help": '<rect x="2" y="2" width="12" height="12" fill="#2f6bb0" stroke="#184a82"/><text x="8" y="11.5" text-anchor="middle" font-family="Tahoma" font-size="9" font-weight="bold" fill="#fff">?</text>',
}


def data_uri(svg: str) -> str:
    return "data:image/svg+xml," + quote(svg, safe="")


def build_icon_css() -> str:
    rules = [
        ".tb{background-repeat:no-repeat;background-position:center;background-size:16px 16px}",
        f'.word-icon,.task-icon{{background-image:url("{data_uri(WORD_ICON_SVG)}");background-repeat:no-repeat;background-position:center;background-size:16px 16px}}',
        f'.start-icon{{width:15px;height:15px;background-image:url("{data_uri(START_ICON_SVG)}");background-repeat:no-repeat;background-position:center;background-size:15px 15px}}',
        f'.help-q{{width:13px;height:13px;background-image:url("{data_uri(svg_16(ICONS["help"]))}");background-repeat:no-repeat;background-position:center;background-size:13px 13px}}',
        '.clippy-figure{background-image:url("app/static/clippy.png");background-repeat:no-repeat;background-position:center bottom;background-size:contain}',
    ]
    for name, inner in ICONS.items():
        uri = data_uri(svg_16(inner))
        rules.append(f'.tb-{name}{{background-image:url("{uri}")}}')

    extras = {
        "ql-doc": svg_16(
            '<path d="M3 1.5h7L13 4.5V14H3z" fill="#fff" stroke="#dcdcdc"/>'
            '<path d="M10 1.5v3h3" fill="#cfe0ff" stroke="#9db6d8"/>'
            '<path d="M5 7h6M5 9h6M5 11h4" stroke="#2f6bb0"/>'
        ),
        "ql-folder": svg_16(
            '<path d="M1.5 4h5l1 1.5h7V13h-13z" fill="#f3d27a" stroke="#7a5b12"/>'
            '<path d="M1.5 6.5h13" stroke="#7a5b12"/>'
        ),
        "ql-desktop": svg_16(
            '<rect x="1.5" y="2.5" width="13" height="9" fill="#dfe9f5" stroke="#4c5a6b"/>'
            '<rect x="3" y="4" width="10" height="6" fill="#3b7ec9"/>'
            '<path d="M5 13.5h6M8 11.5v2" stroke="#4c5a6b"/>'
        ),
        "tray-vol": svg_16(
            '<path d="M2 6h3l3-3v10l-3-3H2z" fill="#d9e6f7" stroke="#0b3a72"/>'
            '<path d="M10 5.5c1.4 1.2 1.4 3.8 0 5M12 3.5c2.4 2.2 2.4 6.8 0 9" fill="none" stroke="#d9e6f7" stroke-width="1.3"/>'
        ),
        "tray-net": svg_16(
            '<rect x="1.5" y="3" width="10" height="7" fill="#d9e6f7" stroke="#0b3a72"/>'
            '<path d="M5 13h5M8 10v3" stroke="#0b3a72"/>'
            '<rect x="11.5" y="8" width="3" height="5" fill="#bfe0ad" stroke="#0b3a72"/>'
        ),
        "sm-github": svg_16(
            '<circle cx="8" cy="8" r="5.6" fill="#e8f0fb" stroke="#173d7a"/>'
            '<path d="M8 2.4c-3.1 0-5.6 2.5-5.6 5.6 0 2.5 1.6 4.6 3.8 5.3.3.1.4-.1.4-.3v-1.1c-1.6.3-1.9-.7-1.9-.7-.3-.7-.6-.9-.6-.9-.5-.3 0-.3 0-.3.6 0 .9.6.9.6.5.9 1.4.6 1.7.5.1-.4.2-.6.4-.8-1.3-.1-2.6-.6-2.6-2.8 0-.6.2-1.1.6-1.5-.1-.1-.2-.7.1-1.4 0 0 .5-.2 1.6.6.5-.1 1-.2 1.5-.2s1 0 1.5.2c1.1-.8 1.6-.6 1.6-.6.3.7.1 1.3.1 1.4.4.4.6.9.6 1.5 0 2.2-1.3 2.7-2.6 2.8.2.2.4.6.4 1.2v1.8c0 .2.1.4.4.3 2.2-.7 3.8-2.8 3.8-5.3 0-3.1-2.5-5.6-5.6-5.6z" fill="#173d7a"/>'
        ),
        "sm-print": svg_16(ICONS["print"]),
        "sm-map": svg_16(
            '<rect x="2" y="2.5" width="12" height="11" fill="#fff" stroke="#4a5f7f"/>'
            '<path d="M4.5 5.5h7M4.5 8h7M4.5 10.5h4.5" stroke="#2f6bb0"/>'
        ),
        "sm-full": svg_16(
            '<rect x="2" y="2.5" width="12" height="11" fill="#dbe8f8" stroke="#4a5f7f"/>'
            '<path d="M4.5 5h7v6h-7z" fill="#fff" stroke="#2f6bb0"/>'
            '<path d="M1.5 1.5h3M11.5 1.5h3M1.5 14.5h3M11.5 14.5h3" stroke="#173d7a"/>'
        ),
        "sm-help": svg_16(ICONS["help"]),
        "sm-info": svg_16(
            '<circle cx="8" cy="8" r="6" fill="#2f6bb0" stroke="#184a82"/>'
            '<path d="M8 7v4M8 4.8v.2" stroke="#fff" stroke-width="1.6"/>'
        ),
        "sm-power": svg_16(
            '<circle cx="8" cy="8.4" r="5.2" fill="#f6d6d0" stroke="#8b1a0b"/>'
            '<path d="M8 3v5" stroke="#c22" stroke-width="1.8"/>'
            '<path d="M5.2 5.2a4 4 0 1 0 5.6 0" fill="none" stroke="#c22" stroke-width="1.6"/>'
        ),
    }
    for name, svg in extras.items():
        uri = data_uri(svg)
        rules.append(
            f'.{name}{{background-image:url("{uri}");background-repeat:no-repeat;'
            'background-position:center;background-size:16px 16px}'
        )
    rules.append(".tray-icon{background-size:14px 14px}")
    return "\n".join(rules)


def icon(name: str, label: str, extra: str = "") -> str:
    classes = f"tb tb-{name}" + (f" {extra}" if extra else "")
    return f'<span class="{classes}" title="{label}" aria-hidden="true"></span>'


SEP = '<span class="tb-sep" aria-hidden="true"></span>'


def tb_button(action: str, name: str, label: str, state: str = "", extra: str = "") -> str:
    state_attr = f' data-state-command="{state}"' if state else ""
    classes = f"tb tb-{name}" + (f" {extra}" if extra else "")
    return (
        f'<button type="button" class="{classes}" data-action="{action}"{state_attr} '
        f'title="{label}" aria-label="{label}"></button>'
    )


def fmt_button(
    action: str, text: str, label: str, css_class: str, state: str
) -> str:
    return (
        f'<button type="button" class="fmt-btn {css_class}" data-action="{action}" '
        f'data-state-command="{state}" title="{label}" aria-label="{label}" '
        f'aria-pressed="false">{text}</button>'
    )


def combo_button(key: str, label: str, css_class: str) -> str:
    return (
        f'<button type="button" class="combo {css_class}" data-combo="{key}" '
        f'aria-haspopup="listbox" aria-expanded="false" title="{label}">'
        f'<span class="combo-label">{label}</span><span class="dd" aria-hidden="true">▾</span></button>'
    )


def menu_item(
    label: str,
    action: str = "",
    shortcut: str = "",
    check: str = "",
    table: bool = False,
    disabled: bool = False,
) -> str:
    if label == "-":
        return '<div class="menu-sep" role="separator"></div>'
    attrs = ['type="button"', 'class="menu-item"']
    if action:
        attrs.append(f'data-action="{action}"')
    if check:
        attrs.append(f'data-check="{check}" aria-checked="false" role="menuitemcheckbox"')
    else:
        attrs.append('role="menuitem"')
    if table:
        attrs.append("data-requires-table")
    if disabled:
        attrs.append("disabled")
    shortcut_html = (
        f'<span class="mi-shortcut" aria-hidden="true">{shortcut}</span>'
        if shortcut
        else ""
    )
    return (
        f'<button {" ".join(attrs)}>'
        '<span class="mi-check" aria-hidden="true"></span>'
        f'<span class="mi-label">{label}</span>{shortcut_html}</button>'
    )


MENUS = [
    (
        "file",
        "File",
        [
            menu_item("New…", "file-new", "Ctrl+N"),
            menu_item("Open…", "file-open", "Ctrl+O"),
            menu_item("Save", "file-save", "Ctrl+S"),
            menu_item("Save As…", "file-save-as", "F12"),
            menu_item("-"),
            menu_item("Print…", "file-print", "Ctrl+P"),
            menu_item("Print Preview", "file-print", "Ctrl+F2"),
            menu_item("-"),
            menu_item("Properties", "file-properties"),
            menu_item("-"),
            menu_item("Revert to Portfolio", "file-revert"),
            menu_item("-"),
            menu_item("Exit", "file-exit"),
        ],
    ),
    (
        "edit",
        "Edit",
        [
            menu_item("Undo", "edit-undo", "Ctrl+Z"),
            menu_item("Redo", "edit-redo", "Ctrl+Y"),
            menu_item("-"),
            menu_item("Cut", "edit-cut", "Ctrl+X"),
            menu_item("Copy", "edit-copy", "Ctrl+C"),
            menu_item("Paste", "edit-paste", "Ctrl+V"),
            menu_item("-"),
            menu_item("Select All", "edit-select-all", "Ctrl+A"),
            menu_item("Find…", "edit-find", "Ctrl+F"),
            menu_item("-"),
            menu_item("Office Clipboard…", "edit-clipboard"),
        ],
    ),
    (
        "view",
        "View",
        [
            menu_item("Document Map", "view-docmap", check="docmap"),
            menu_item("Ruler", "view-ruler", check="ruler"),
            menu_item("-"),
            menu_item("Standard Toolbar", "view-standard", check="standard-toolbar"),
            menu_item("Formatting Toolbar", "view-format", check="formatting-toolbar"),
            menu_item("-"),
            menu_item("Zoom…", "view-zoom"),
            menu_item("Full Screen", "view-fullscreen", "F11"),
        ],
    ),
    (
        "insert",
        "Insert",
        [
            menu_item("Page Break", "insert-page-break", "Ctrl+Enter"),
            menu_item("Date and Time…", "insert-date"),
            menu_item("Horizontal Line", "insert-rule"),
            menu_item("-"),
            menu_item("Hyperlink…", "insert-hyperlink", "Ctrl+K"),
            menu_item("Table…", "insert-table"),
        ],
    ),
    (
        "format",
        "Format",
        [
            menu_item("Font…", "format-font"),
            menu_item("Style…", "format-style"),
            menu_item("-"),
            menu_item("Bold", "format-bold", "Ctrl+B"),
            menu_item("Italic", "format-italic", "Ctrl+I"),
            menu_item("Underline", "format-underline", "Ctrl+U"),
            menu_item("-"),
            menu_item("Bullets and Numbering…", "format-bullets"),
            menu_item("Borders and Shading…", "format-border"),
            menu_item("-"),
            menu_item("Increase Indent", "format-indent"),
        ],
    ),
    (
        "tools",
        "Tools",
        [
            menu_item("Spelling and Grammar…", "tools-spelling", "F7"),
            menu_item("Word Count…", "tools-word-count"),
            menu_item("-"),
            menu_item("Options…", "tools-options"),
            menu_item("-"),
            menu_item("GitHub Repositories", "help-contact"),
        ],
    ),
    (
        "table",
        "Table",
        [
            menu_item("Insert Table…", "insert-table"),
            menu_item("-"),
            menu_item("Insert Row Above", "table-row-above", table=True),
            menu_item("Insert Row Below", "table-row-below", table=True),
            menu_item("Insert Column Left", "table-col-left", table=True),
            menu_item("Insert Column Right", "table-col-right", table=True),
            menu_item("-"),
            menu_item("Delete Row", "table-delete-row", table=True),
            menu_item("Delete Column", "table-delete-col", table=True),
            menu_item("-"),
            menu_item("Select Table", "table-select", table=True),
        ],
    ),
    (
        "window",
        "Window",
        [
            menu_item("Portfolio.doc", "focus-document", check="window-document"),
            menu_item("-"),
            menu_item("Minimize", "window-minimize"),
            menu_item("Restore", "window-restore"),
            menu_item("Close", "window-close"),
        ],
    ),
    (
        "help",
        "Help",
        [
            menu_item("Portfolio Help", "help-topics", "F1"),
            menu_item("Show the Assistant", "help-assistant"),
            menu_item("Keyboard Shortcuts", "help-shortcuts"),
            menu_item("-"),
            menu_item("About Portfolio.doc", "help-about"),
            menu_item("About Microsoft Word", "help-about-word"),
            menu_item("-"),
            menu_item("Contact the Author", "help-contact"),
        ],
    ),
]


def render_menus() -> str:
    chunks = []
    for key, label, items in MENUS:
        chunks.append(
            f'<div class="menu" data-menu-name="{key}">'
            f'<button type="button" class="menu-title" data-menu="{key}" '
            f'aria-haspopup="true" aria-expanded="false">{label}</button>'
            f'<div class="menu-drop" role="menu" aria-label="{label}">'
            + "".join(items)
            + "</div></div>"
        )
    return "".join(chunks)


def project_row(project: dict) -> str:
    stack = " · ".join(project["tags"])
    return f"""
        <tr>
          <td class="col-project">
            <a class="proj-name" href="{project['url']}" target="_blank" rel="noreferrer">{project['title']}</a>
            <div class="proj-desc">{project['description']}</div>
            <div class="proj-stack-inline">Stack: {stack}</div>
          </td>
          <td class="col-stack">{stack}</td>
          <td class="col-link"><a href="{project['url']}" target="_blank" rel="noreferrer">Source ↗</a></td>
        </tr>"""


def build_page() -> str:
    standard_tools = (
        tb_button("file-new", "new", "New (Ctrl+N)")
        + tb_button("file-open", "open", "Open (Ctrl+O)")
        + tb_button("file-save", "save", "Save (Ctrl+S)")
        + tb_button("file-print", "print", "Print (Ctrl+P)")
        + SEP
        + tb_button("edit-undo", "undo", "Undo (Ctrl+Z)")
        + tb_button("edit-redo", "redo", "Redo (Ctrl+Y)")
        + SEP
        + tb_button("insert-hyperlink", "link", "Insert Hyperlink (Ctrl+K)")
        + tb_button("insert-table", "table", "Insert Table")
        + tb_button("tools-spelling", "spell", "Spelling and Grammar (F7)")
        + SEP
        + tb_button("view-zoom", "zoom", "Zoom")
    )
    rows = "".join(project_row(project) for project in PROJECTS)
    now = datetime.now()
    year = date.today().year
    clock = now.strftime("%I:%M %p").lstrip("0")

    return f"""
<div class="xp-desktop">
  <div class="word-window" role="region" aria-label="Portfolio document in a Microsoft Word window">

    <div class="titlebar" id="word-titlebar">
      <span class="word-icon" aria-hidden="true"></span>
      <span class="title-text">Portfolio.doc - Microsoft Word</span>
      <span class="window-controls">
        <button type="button" class="wc wc-min" data-action="window-minimize" title="Minimize" aria-label="Minimize"></button>
        <button type="button" class="wc wc-max" data-action="window-restore" title="Restore Down" aria-label="Restore Down"></button>
        <button type="button" class="wc wc-close" data-action="window-close" title="Close" aria-label="Close"></button>
      </span>
    </div>

    <div class="menubar" role="menubar" aria-label="Application menu">
      {render_menus()}
    </div>

    <div class="toolbar toolbar-standard" role="toolbar" aria-label="Standard toolbar">
      {standard_tools}
      <span class="toolbar-spacer"></span>
      <button type="button" class="help-box" data-action="help-topics" title="Type a question for help">
        <span class="help-q" aria-hidden="true"></span>Type a question for help<span class="help-caret" aria-hidden="true">▾</span>
      </button>
    </div>

    <div class="toolbar toolbar-format" role="toolbar" aria-label="Formatting toolbar">
      {combo_button("style", "Normal", "combo-style")}
      {SEP}
      {combo_button("font", "Times New Roman", "combo-font")}
      {combo_button("size", "12", "combo-size")}
      {SEP}
      {fmt_button("format-bold", "B", "Bold (Ctrl+B)", "fmt-bold", "bold")}
      {fmt_button("format-italic", "I", "Italic (Ctrl+I)", "fmt-italic", "italic")}
      {fmt_button("format-underline", "U", "Underline (Ctrl+U)", "fmt-underline", "underline")}
      {SEP}
      {tb_button("align-left", "align-left", "Align Left", "justifyLeft")}
      {tb_button("align-center", "align-center", "Center", "justifyCenter")}
      {tb_button("align-right", "align-right", "Align Right", "justifyRight")}
      {tb_button("align-justify", "align-justify", "Justify", "justifyFull")}
      {SEP}
      {tb_button("format-bullets", "bullets", "Bulleted List", "insertUnorderedList", "fmt-extra")}
      {tb_button("format-numbers", "numbers", "Numbered List", "insertOrderedList", "fmt-extra")}
      {tb_button("format-indent", "indent", "Increase Indent", "indent", "fmt-extra")}
      {SEP}
      {tb_button("format-border", "border", "Outside Border", "", "fmt-extra")}
      {tb_button("format-highlight", "highlight", "Highlight", "", "fmt-extra")}
      {tb_button("format-font-color", "font-color", "Font Color")}
    </div>

    <nav class="mobile-nav" aria-label="Portfolio sections">
      <a href="#profile">Profile</a>
      <a href="#work">Projects</a>
      <a href="#toolkit">Toolkit</a>
      <a href="#contact">Contact</a>
      <a href="{GITHUB_URL}" target="_blank" rel="noreferrer">GitHub ↗</a>
    </nav>

    <div class="window-body">
      <nav class="docmap" id="docmap" aria-label="Document map">
        <div class="docmap-head"><span>Document Map</span><button type="button" class="docmap-x" data-action="view-docmap" title="Close Document Map" aria-label="Close Document Map">×</button></div>
        <button type="button" class="docmap-filter" data-combo="docmap" aria-haspopup="listbox" aria-expanded="false" title="Document Map view">
          <span class="combo-label">Headings</span><span class="dd" aria-hidden="true">▾</span>
        </button>
        <div class="docmap-tree" data-docmap-view="headings">
          <a class="lvl1" href="#top">Gabriel Elorde</a>
          <a class="lvl2" href="#profile">Profile</a>
          <a class="lvl2" href="#work">Selected Projects</a>
          <a class="lvl3" href="#featured">Artifact Speculation Engine</a>
          <a class="lvl3" href="#projects">Project List</a>
          <a class="lvl2" href="#education">Education</a>
          <a class="lvl2" href="#toolkit">Toolkit</a>
          <a class="lvl2" href="#contact">Contact</a>
        </div>
        <div class="docmap-tree" data-docmap-view="pages" hidden>
          <a class="lvl1" href="#top">Page 1</a>
          <a class="lvl2" href="#education">Page 2</a>
        </div>
      </nav>

      <div class="editor">
        <div class="ruler-shell" id="ruler-shell">
          <div class="ruler" aria-hidden="true">
            <span class="ruler-margin ruler-margin-l"></span>
            <span class="ruler-margin ruler-margin-r"></span>
            <span class="ruler-track">
              <span class="ruler-num" style="left:48px">1</span>
              <span class="ruler-num" style="left:144px">2</span>
              <span class="ruler-num" style="left:240px">3</span>
              <span class="ruler-num" style="left:336px">4</span>
              <span class="ruler-num" style="left:432px">5</span>
              <span class="ruler-num" style="left:528px">6</span>
            </span>
            <span class="indent-mark indent-first"></span>
            <span class="indent-mark indent-left"></span>
            <span class="indent-mark indent-right"></span>
          </div>
        </div>

        <div class="editor-scroll" id="editor-scroll">
          <div class="paper-stack" id="paper-stack" contenteditable="true" spellcheck="true" role="textbox" aria-multiline="true" aria-label="Portfolio document. Click to edit, select text to format, and use the menus or toolbar.">
            <article class="sheet" id="top">
              <header class="doc-head">
                <p class="doc-eyebrow">Curriculum Vitae · {year}</p>
                <h1 class="doc-name">Gabriel Elorde</h1>
                <p class="doc-role">Computer Science Student · Software Developer</p>
                <p class="doc-meta">
                  Cebu, Philippines
                  <span class="meta-sep">|</span>
                  <a href="{GITHUB_URL}" target="_blank" rel="noreferrer">github.com/maidenator</a>
                  <span class="meta-sep">|</span>
                  <a href="#contact">Contact details</a>
                </p>
              </header>
              <hr class="doc-rule" />

              <section class="doc-section" id="profile">
                <h2 class="doc-h2">Profile</h2>
                <p>
                  I'm Gabriel, a third-year Computer Science student at
                  <strong>Cebu Institute of Technology University</strong>. My
                  GitHub has coursework, side projects, and other experiments.
                </p>
                <p>
                  Lately I've been working on the <strong>Artifact Speculation Engine</strong>,
                  a Monte Carlo simulator with a C++/WebAssembly core. I've also
                  been solving algorithm problems, working on browser extensions,
                  and building neural networks from scratch.
                </p>
              </section>

              <section class="doc-section" id="work">
                <h2 class="doc-h2">Selected Projects</h2>

                <div class="callout" id="featured">
                  <p class="callout-label">Featured Project</p>
                  <h3 class="callout-title">Artifact Speculation Engine</h3>
                  <p class="callout-desc">
                    A browser-based simulator for Genshin Impact artifact farming.
                    A C++ engine compiled to WebAssembly runs Monte Carlo trials,
                    and a React front end shows your success odds, percentiles,
                    and the resin you'll likely spend.
                  </p>
                  <p class="callout-stack">
                    <span class="label">Stack:</span> TypeScript · React · C++ ·
                    WebAssembly · Monte Carlo
                  </p>
                  <p class="callout-links">
                    <a href="{FEATURED_LIVE_URL}" target="_blank" rel="noreferrer">View live demo ↗</a>
                    <span class="link-sep">·</span>
                    <a href="{FEATURED_REPO_URL}" target="_blank" rel="noreferrer">Source code ↗</a>
                  </p>
                </div>

                <p class="doc-note">
                  Selected public repositories. Each title links to the source
                  on GitHub.
                </p>

                <table class="doc-table projects-table" id="projects">
                  <caption class="sr-only">Selected projects with stacks and repository links</caption>
                  <thead>
                    <tr>
                      <th class="col-project">Project</th>
                      <th class="col-stack">Stack</th>
                      <th class="col-link"><span class="th-wide">Repository</span><span class="th-narrow">Link</span></th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows}
                  </tbody>
                </table>

                <p class="table-note">
                  Browse all <a href="{REPOSITORIES_URL}" target="_blank" rel="noreferrer">{PUBLIC_REPO_COUNT} public repositories on GitHub ↗</a>
                </p>
              </section>

              <footer class="sheet-footer">
                <span>Gabriel Elorde · Curriculum Vitae</span>
                <span>Page 1 of 2</span>
              </footer>
            </article>

            <article class="sheet sheet-two">
              <section class="doc-section" id="education">
                <h2 class="doc-h2">Education</h2>
                <table class="doc-table edu-table">
                  <tbody>
                    <tr>
                      <td class="edu-school">
                        <strong>Cebu Institute of Technology University</strong>
                        <div>BS Computer Science · third year</div>
                      </td>
                      <td class="edu-loc">Cebu, Philippines</td>
                    </tr>
                  </tbody>
                </table>
                <p class="table-note">
                  Coursework so far includes design and analysis of algorithms,
                  data structures, and software development.
                </p>
              </section>

              <section class="doc-section" id="toolkit">
                <h2 class="doc-h2">Toolkit</h2>
                <table class="doc-table toolkit-table">
                  <thead>
                    <tr><th>Area</th><th>Tools</th></tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td class="tk-area">Languages</td>
                      <td>TypeScript · JavaScript · C++ · C# · Java · Kotlin · PHP · Python</td>
                    </tr>
                    <tr>
                      <td class="tk-area">Web &amp; UI</td>
                      <td>React · Blazor · Streamlit · HTML &amp; CSS</td>
                    </tr>
                    <tr>
                      <td class="tk-area">Tools &amp; ideas</td>
                      <td>WebAssembly · Git · Monte Carlo · Neural networks</td>
                    </tr>
                  </tbody>
                </table>
              </section>

              <section class="doc-section contact-section" id="contact">
                <h2 class="doc-h2">Contact</h2>
                <table class="doc-table contact-table">
                  <tbody>
                    <tr><td class="ct-key">GitHub</td><td><a href="{GITHUB_URL}" target="_blank" rel="noreferrer">github.com/maidenator</a></td></tr>
                    <tr><td class="ct-key">Location</td><td>Cebu, Philippines</td></tr>
                    <tr><td class="ct-key">Studying</td><td>BS Computer Science, CIT-U</td></tr>
                    <tr><td class="ct-key">Currently</td><td><a href="#featured">Building the Artifact Speculation Engine</a></td></tr>
                  </tbody>
                </table>
                <p class="contact-action">
                  <a class="doc-button" href="{GITHUB_URL}" target="_blank" rel="noreferrer">Find me on GitHub →</a>
                </p>
              </section>

              <footer class="sheet-footer">
                <span>Gabriel Elorde · Curriculum Vitae</span>
                <span>Page 2 of 2</span>
              </footer>
            </article>
          </div>
        </div>
      </div>
    </div>

    <div class="statusbar">
      <button type="button" class="sb sb-page sb-action" id="status-page" data-action="go-to-page" title="Go to page">Page 1 of 2</button>
      <span class="sb sb-optional">Sec 1</span>
      <span class="sb sb-optional">1/1</span>
      <span class="sb sb-optional sb-pos" id="status-pos">At 1.0"</span>
      <span class="sb sb-optional" id="status-line">Ln 1</span>
      <span class="sb sb-optional" id="status-col">Col 1</span>
      <span class="sb sb-grow"></span>
      <span class="sb sb-toggle sb-optional">REC</span>
      <span class="sb sb-toggle sb-optional">TRK</span>
      <span class="sb sb-toggle sb-optional">EXT</span>
      <span class="sb sb-toggle sb-optional">OVR</span>
      <button type="button" class="sb sb-zoom sb-action" id="status-zoom" data-action="view-zoom" title="Zoom">100%</button>
      <button type="button" class="sb sb-lang sb-action" id="status-lang" data-action="tools-language" title="Set Language">English (U.S.)</button>
    </div>
  </div>

  <div class="taskbar" id="taskbar">
    <button type="button" class="start-btn" id="start-button" aria-haspopup="true" aria-expanded="false" title="Start">
      <span class="start-icon" aria-hidden="true"></span>
      <span>start</span>
    </button>
    <span class="quick-launch">
      <button type="button" class="ql ql-doc" data-action="focus-document" title="Portfolio.doc" aria-label="Open Portfolio.doc"></button>
      <a class="ql ql-folder" href="{REPOSITORIES_URL}" target="_blank" rel="noreferrer" title="Open GitHub repositories" aria-label="Open GitHub repositories"></a>
      <button type="button" class="ql ql-desktop" data-action="show-desktop" title="Show Desktop" aria-label="Show Desktop"></button>
    </span>
    <button type="button" class="task-btn is-active" id="task-button" data-action="task-toggle" aria-pressed="true" title="Portfolio.doc - Microsoft Word">
      <span class="task-icon" aria-hidden="true"></span>
      <span class="task-label">Portfolio.doc - Microsoft Word</span>
    </button>
    <span class="tray">
      <button type="button" class="tray-icon tray-vol" data-action="tray-volume" title="Volume" aria-label="Volume"></button>
      <button type="button" class="tray-icon tray-net" data-action="tray-network" title="Local Area Connection" aria-label="Local Area Connection"></button>
      <button type="button" class="clock" id="clock" data-action="clock" title="{now.strftime('%A, %B %d, %Y')}">{clock}</button>
    </span>
  </div>

  <div class="start-menu" id="start-menu" hidden>
    <div class="sm-header">
      <span class="sm-avatar" aria-hidden="true">GE</span>
      <span class="sm-user"><strong>Gabriel Elorde</strong><small>Portfolio session</small></span>
    </div>
    <div class="sm-columns">
      <div class="sm-col">
        <button type="button" class="sm-item" data-action="focus-document">
          <span class="sm-icon ql-doc" aria-hidden="true"></span>
          <span><strong>Portfolio.doc</strong><small>Microsoft Word document</small></span>
        </button>
        <button type="button" class="sm-item" data-action="file-open">
          <span class="sm-icon ql-folder" aria-hidden="true"></span>
          <span><strong>My Documents</strong><small>Open a document</small></span>
        </button>
        <button type="button" class="sm-item" data-action="help-contact">
          <span class="sm-icon sm-github" aria-hidden="true"></span>
          <span><strong>GitHub Repositories</strong><small>github.com/maidenator</small></span>
        </button>
        <button type="button" class="sm-item" data-action="file-print">
          <span class="sm-icon sm-print" aria-hidden="true"></span>
          <span><strong>Print Portfolio</strong><small>Two-page curriculum vitae</small></span>
        </button>
      </div>
      <div class="sm-col sm-col-secondary">
        <button type="button" class="sm-item" data-action="view-docmap">
          <span class="sm-icon sm-map" aria-hidden="true"></span><span><strong>Document Map</strong></span>
        </button>
        <button type="button" class="sm-item" data-action="view-fullscreen">
          <span class="sm-icon sm-full" aria-hidden="true"></span><span><strong>Full Screen</strong></span>
        </button>
        <button type="button" class="sm-item" data-action="help-topics">
          <span class="sm-icon sm-help" aria-hidden="true"></span><span><strong>Portfolio Help</strong></span>
        </button>
        <button type="button" class="sm-item" data-action="help-about">
          <span class="sm-icon sm-info" aria-hidden="true"></span><span><strong>About Portfolio.doc</strong></span>
        </button>
      </div>
    </div>
    <div class="sm-footer">
      <button type="button" data-action="start-shutdown">
        <span class="sm-icon sm-power" aria-hidden="true"></span>Turn Off Computer
      </button>
    </div>
  </div>

  <details class="clippy" id="clippy">
    <summary class="clippy-summary" title="Need help? Click the paperclip." aria-label="Toggle the portfolio assistant">
      <span class="clippy-figure" aria-hidden="true"></span>
    </summary>
    <div class="clippy-bubble">
      <strong>It looks like you're editing a portfolio.</strong>
      Try the menus, select text to format, or press Ctrl+P to print.
      <span class="clippy-links">
        <a href="#work">Show projects →</a><span>·</span>
        <button type="button" class="clippy-more" data-action="help-topics">What works? →</button>
      </span>
      <span class="clippy-hint">Click the paperclip to hide.</span>
    </div>
  </details>

  <div class="desktop-reopen" id="desktop-reopen" hidden>
    <button type="button" class="desktop-shortcut" data-action="reopen" title="Drag to move, click to reopen Portfolio.doc">
      <span class="desktop-doc-icon ql-doc" aria-hidden="true"></span>
      <span>Portfolio.doc</span>
    </button>
    <p class="desktop-hint">Drag the icon to move it, or click to reopen your portfolio.</p>
  </div>

  <div class="dialog-layer" id="dialog-layer" hidden>
    <div class="word-dialog" id="word-dialog" role="dialog" aria-modal="true" aria-labelledby="dialog-title">
      <div class="dialog-titlebar" id="dialog-titlebar">
        <span class="word-icon dialog-word-icon" aria-hidden="true"></span>
        <span class="dialog-title-text" id="dialog-title">Microsoft Word</span>
        <button type="button" class="dialog-close" data-dialog-close aria-label="Close dialog">×</button>
      </div>
      <div class="dialog-body" id="dialog-body"></div>
      <div class="dialog-buttons" id="dialog-buttons"></div>
    </div>
  </div>

  <div class="shutdown-screen" id="shutdown-screen" hidden>
    <div class="shutdown-inner">
      <p class="shutdown-message" id="shutdown-message">It's now safe to turn off your computer.</p>
      <button type="button" class="shutdown-restart" data-action="restart-session">Restart</button>
    </div>
  </div>

  <div class="word-toast" id="word-toast" role="status" aria-live="polite" hidden></div>
  <input type="file" id="open-file" accept=".txt,.html,.htm,.doc" hidden />
  <div class="combo-popup" id="combo-popup" role="listbox" hidden></div>
</div>
    """


css = (BASE_DIR / "styles.css").read_text(encoding="utf-8")
js_path = BASE_DIR / "static" / "word-ui.js"
js_version = int(js_path.stat().st_mtime)
st.html(
    f"<style>{css}{build_icon_css()}</style>{build_page()}"
    f'<script src="app/static/word-ui.js?v={js_version}"></script>',
    unsafe_allow_javascript=True,
)
