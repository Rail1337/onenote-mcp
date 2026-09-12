"""
OneNote MCP Server
===================
An MCP (Model Context Protocol) server that reads local OneNote (.one) files
directly from disk and reads/writes/organizes OneNote via the COM API.
No Azure registration or authentication needed.

Reading: parses local backup files, auto-detected regardless of Windows/
Office display language (e.g. English "Backup", German "Sicherung").

Live: talks to the running OneNote desktop app via its COM API (PowerShell
under the hood) for always-current reads, writing, and organizing.

It exposes tools for Claude to:
    - List, read, and search notebooks/sections from the local backup
    - Read a single page or a whole section live from the running app
    - Create, rename, move, and delete sections and section groups
    - Create pages, append to pages, and rename pages
    - List items in the OneNote recycle bin, and restore them via move
    (deletes require an explicit confirm=true and only move items to the
    recycle bin, never permanent)
    - Find/replace text on a page, replace its last block, or append to it
    - List recent actions, undo the last one, and redo the last undo --
    backed by a persistent history.md log that survives restarts (its
    on-disk path can be looked up directly, to read or archive-check it)
    - Insert an image from a local file, and list what's in the drop folder

Prerequisites:
    pip install "mcp[cli]" pyOneNote
    (or: uv add "mcp[cli]" pyOneNote)
    + OneNote desktop app (for live/write features)

Usage with Claude Code:
    claude mcp add --transport stdio onenote -- uv --directory "path/to/this/project" run server.py
"""

import base64
import html
import io
import json
import logging
import os
import re
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from pyOneNote.OneDocument import OneDocment
from mcp.server.fastmcp import FastMCP
from PIL import Image as PILImage

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Where OneNote stores local backup files. The folder is auto-detected because
# its name is localized by Windows/Office language (English: "Backup", German:
# "Sicherung", French: "Sauvegarde", ...). Override with ONENOTE_BACKUP_DIR if
# your backups live somewhere non-standard.
_NON_BACKUP_DIR_NAMES = {
    "accessibilitycheckerindex", "fulltextsearchindex", "masterindex",
    "serverlistings", "cache", "offlinefilesinfo",
}
_KNOWN_BACKUP_DIR_NAMES = (
    "Backup", "Sicherung", "Sauvegarde", "Copia di backup",
    "Copia de seguridad", "Reservekopie", "Backup-kopie",
)


def _detect_backup_dir() -> Path:
    env_override = os.environ.get("ONENOTE_BACKUP_DIR")
    if env_override:
        return Path(env_override)

    onenote_root = Path(
        os.environ.get("LOCALAPPDATA", "")
    ) / "Microsoft" / "OneNote" / "16.0"

    if onenote_root.is_dir():
        # Fast path: try known localized names first.
        for candidate_name in _KNOWN_BACKUP_DIR_NAMES:
            candidate = onenote_root / candidate_name
            if candidate.is_dir() and any(candidate.rglob("*.one")):
                return candidate

        # Fallback: scan every subfolder and use whichever one actually
        # contains .one files (covers languages not in the list above).
        for entry in onenote_root.iterdir():
            if not entry.is_dir():
                continue
            if entry.name.lower() in _NON_BACKUP_DIR_NAMES:
                continue
            if any(entry.rglob("*.one")):
                return entry

    # Nothing found; return the English default so the startup error message
    # points at a sensible path.
    return onenote_root / "Backup"


ONENOTE_DIR = _detect_backup_dir()

def _get_windows_downloads_folder() -> Path | None:
    """Ask Windows where the Downloads folder actually is via
    SHGetKnownFolderPath, instead of assuming the default
    %USERPROFILE%\\Downloads -- Windows lets users relocate Downloads (and
    Documents, Pictures, etc.) to another drive entirely via
    Properties > Location, which Path.home()/"Downloads" can't see."""
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        # FOLDERID_Downloads -- this GUID has no CSIDL equivalent, it only
        # exists via the newer SHGetKnownFolderPath API.
        FOLDERID_Downloads = GUID(
            0x374DE290, 0x123F, 0x4565,
            (ctypes.c_ubyte * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B),
        )

        path_ptr = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(FOLDERID_Downloads), 0, None, ctypes.byref(path_ptr)
        )
        if result == 0 and path_ptr.value:
            path = Path(path_ptr.value)
            ctypes.windll.ole32.CoTaskMemFree(path_ptr)
            return path
    except Exception as e:
        # Logging isn't configured yet at this point in module load, so this
        # goes straight to stderr instead of through the `log` object.
        print(f"Could not resolve the real Downloads folder via Windows: {e}", file=sys.stderr)
    return None


# Default folder to look for images in when insert_image_from_file is given
# just a filename instead of a full path. Override with ONENOTE_IMAGE_DIR to
# use a different drop folder.
IMAGE_DROP_DIR = Path(
    os.environ.get("ONENOTE_IMAGE_DIR")
    or str(_get_windows_downloads_folder() or (Path.home() / "Downloads"))
)

# ---------------------------------------------------------------------------
# Action history (persisted to a markdown file so it survives restarts and
# can be reviewed by the user directly, or by Claude in a later session).
# Not committed to the repo -- this is personal edit history, see .gitignore.
# ---------------------------------------------------------------------------

HISTORY_FILE = Path(__file__).resolve().parent / "history.md"
HISTORY_ARCHIVE_FILE = Path(__file__).resolve().parent / "history.archive.md"
HISTORY_LIMIT = 500


def _local_timestamp() -> str:
    """Return the current time as an ISO-ish string in the machine's own
    local timezone (e.g. Europe/Berlin), not UTC -- datetime.now() with no
    tzinfo argument already reads the system clock as local wall-clock
    time, so no conversion is needed. history.md is personal, read
    directly by the user, not a cross-timezone API log, so showing the
    time they'd actually see on their own clock is the right default. No
    "Z" suffix, since that specifically denotes UTC and would be wrong here.
    """
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _log_action(summary: str, undo_data: dict) -> None:
    """Append one entry to history.md, then rotate the oldest entries into
    history.archive.md if the active log has grown past HISTORY_LIMIT.

    Each line is "- " followed by one JSON object (timestamp, summary,
    undone, plus whatever undo_last_action needs to reverse the action --
    always including a "type" key). Storing the whole line as JSON, instead
    of human text plus a text-delimited JSON blob, means there's no marker
    string that caller-controlled text (e.g. find_text) could accidentally
    collide with -- JSON's own string escaping handles arbitrary content
    correctly regardless of what it contains. summary is still plain,
    readable text if you open the file directly; list_recent_actions
    re-renders it nicely rather than dumping the raw JSON.
    """
    # _id disambiguates entries for _set_action_undone_flag's lookup --
    # _local_timestamp() only has second precision for readability, so two
    # calls with identical type/fields within the same second produce
    # byte-identical JSON lines; matching by line content alone would then
    # hit whichever occurrence comes first in the file, not the specific
    # one undo/redo actually meant. time.time_ns() effectively never
    # collides.
    entry = {"timestamp": _local_timestamp(), "_id": time.time_ns(), "summary": summary, "undone": False, **undo_data}
    line = f"- {json.dumps(entry)}\n"
    try:
        is_new = not HISTORY_FILE.exists()
        with open(HISTORY_FILE, "a", encoding="utf-8") as f:
            if is_new:
                f.write("# History\n\n")
            f.write(line)
        _rotate_history_if_needed()
    except OSError as e:
        log.warning("Could not write to history log: %s", e)


def _rotate_history_if_needed() -> None:
    """If history.md has grown past HISTORY_LIMIT entries, move the oldest
    overflow entries into history.archive.md. Nothing is ever deleted."""
    try:
        lines = HISTORY_FILE.read_text(encoding="utf-8").splitlines(keepends=True)
    except OSError:
        return

    entry_indices = [i for i, l in enumerate(lines) if l.startswith("- ")]
    if len(entry_indices) <= HISTORY_LIMIT:
        return

    overflow = len(entry_indices) - HISTORY_LIMIT
    cutoff = entry_indices[overflow]  # index of the first entry line to KEEP
    header_lines = lines[: entry_indices[0]]
    archived_lines = lines[entry_indices[0]: cutoff]
    kept_lines = lines[cutoff:]

    try:
        archive_is_new = not HISTORY_ARCHIVE_FILE.exists()
        with open(HISTORY_ARCHIVE_FILE, "a", encoding="utf-8") as f:
            if archive_is_new:
                f.write("# History Archive\n\n")
            f.writelines(archived_lines)
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            f.writelines(header_lines)
            f.writelines(kept_lines)
    except OSError as e:
        log.warning("Could not rotate history log: %s", e)


def _read_history_entries() -> list[str]:
    """Return the raw '- ...' entry lines from history.md, oldest first."""
    try:
        text = HISTORY_FILE.read_text(encoding="utf-8")
    except OSError:
        return []
    return [line for line in text.splitlines() if line.startswith("- ")]


def _parse_history_line(line: str) -> dict:
    """Parse one history entry line (a "- " prefix followed by one JSON
    object) into its timestamp, human summary, undo data, and whether it's
    already marked undone."""
    meta_keys = ("timestamp", "_id", "summary", "undone", "undone_at")
    json_part = line[2:] if line.startswith("- ") else line
    try:
        entry = json.loads(json_part)
    except json.JSONDecodeError as e:
        log.warning("Could not parse history line as JSON: %s", e)
        return {
            "timestamp": "", "id": None, "summary": line.strip(), "undo_data": {},
            "undone": False, "undone_at": None, "raw": line,
        }

    undo_data = {k: v for k, v in entry.items() if k not in meta_keys}
    return {
        "timestamp": entry.get("timestamp", ""),
        "id": entry.get("_id"),
        "summary": entry.get("summary", ""),
        "undo_data": undo_data,
        "undone": bool(entry.get("undone", False)),
        "undone_at": entry.get("undone_at"),
        "raw": line,
    }


def _find_last_undoable_action() -> dict | None:
    """Return the parsed most-recent entry that hasn't been undone yet, or
    None if the log is empty or everything in it is already undone."""
    for line in reversed(_read_history_entries()):
        parsed = _parse_history_line(line)
        if not parsed["undone"]:
            return parsed
    return None


def _find_last_redoable_action() -> dict | None:
    """Return the parsed entry that is currently undone and was undone most
    recently, or None if nothing is in an undone state right now.

    Unlike _find_last_undoable_action, this can't just scan the file in
    reverse: undo_last_action never appends a new line, it flips a flag on
    an existing one, so an *older* entry can end up undone more recently
    than a newer one (e.g. two undo_last_action calls in a row undo the
    last two actions in order, making the second-to-last entry the most
    recently undone). "Most recent" therefore means latest undone_at
    timestamp, not file position.
    """
    candidates = [
        parsed
        for parsed in (_parse_history_line(line) for line in _read_history_entries())
        if parsed["undone"] and parsed["undone_at"]
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda p: p["undone_at"])
    return candidates[-1]


def _set_action_undone_flag(
    entry_id, fallback_raw_line: str, undone: bool, extra: dict | None = None
) -> bool:
    """Rewrite history.md, setting the matching entry's undone flag to
    True (recording when, for _find_last_redoable_action) or back to
    False (on redo, so the entry becomes undoable again -- toggling back
    and forth repeatedly is intentional, there's no separate "redone"
    state).

    Matches primarily by the entry's unique _id (see _log_action) --
    immune to two entries having byte-identical content within the same
    second, which exact-line matching alone is not (_local_timestamp()
    only has second precision, so e.g. two automated find_and_replace
    calls with the same arguments one after another would otherwise be
    indistinguishable, and the wrong one could get flagged). entry_id is
    None for entries logged before _id existed; for those, falls back to
    matching by exact raw line content via fallback_raw_line -- the same
    ambiguity risk older entries always had, not made any worse.

    extra, when given, is merged into the entry at the same time -- used
    by redo_last_action to refresh an entry's object_id/outline_id after
    recreating something, since the recreated object never keeps its
    original ID (see _apply_logged_action's docstring)."""
    try:
        text = HISTORY_FILE.read_text(encoding="utf-8")
    except OSError as e:
        log.warning("Could not read history log to update its undone flag: %s", e)
        return False

    lines = text.splitlines(keepends=True)
    for i, line in enumerate(lines):
        raw = line.rstrip("\n")
        json_part = raw[2:] if raw.startswith("- ") else raw
        try:
            entry = json.loads(json_part)
        except json.JSONDecodeError:
            continue

        matches = (entry.get("_id") == entry_id) if entry_id is not None else (raw == fallback_raw_line)
        if not matches:
            continue

        entry["undone"] = undone
        if undone:
            entry["undone_at"] = _local_timestamp()
        if extra:
            entry.update(extra)
        lines[i] = f"- {json.dumps(entry)}\n"
        try:
            HISTORY_FILE.write_text("".join(lines), encoding="utf-8")
        except OSError as e:
            log.warning("Could not update history log: %s", e)
            return False
        return True
    return False

# ---------------------------------------------------------------------------
# Logging (to stderr so it doesn't break stdio MCP transport)
# ---------------------------------------------------------------------------

LOG_FILE = os.path.join(tempfile.gettempdir(), "onenote_mcp.log")

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger("onenote-mcp")
log.info("Log file: %s", LOG_FILE)

# ---------------------------------------------------------------------------
# OneNote file parsing helpers
# ---------------------------------------------------------------------------


def _discover_notebooks() -> dict[str, dict]:
    """
    Scan the OneNote backup directory and build a notebook → section → files map.

    Returns a dict like:
    {
        "My Notebook": {
            "path": Path(...),
            "sections": {
                "Algorithm": {
                    "files": [Path("Algorithm (On 1-4-2026).one"), ...],
                    "latest": Path(...)   # most recently modified
                },
                ...
            }
        },
        ...
    }
    """
    if not ONENOTE_DIR.exists():
        log.error("OneNote backup directory not found: %s", ONENOTE_DIR)
        return {}

    notebooks = {}
    for notebook_dir in ONENOTE_DIR.iterdir():
        if not notebook_dir.is_dir():
            continue

        notebook_name = notebook_dir.name
        sections: dict[str, dict] = {}

        # Walk all .one files in this notebook (including subdirectories)
        for one_file in notebook_dir.rglob("*.one"):
            # Skip recycle bin
            if "RecycleBin" in str(one_file):
                continue

            # Extract the base section name (strip the date suffix)
            # e.g. "Algorithm (On 1-4-2026).one" → "Algorithm"
            # e.g. "Python.one (On 12-6-2025).one" → "Python"
            fname = one_file.name
            # Remove .one extension(s) and date suffixes
            section_name = re.sub(r"\.one$", "", fname)
            section_name = re.sub(r"\s*\(On \d+-\d+-\d+\)$", "", section_name)
            section_name = re.sub(r"\.one$", "", section_name)  # handle double .one
            section_name = section_name.strip()

            if not section_name:
                section_name = "(unnamed)"

            # Build relative path for context (subfolder within notebook)
            rel_parts = one_file.parent.relative_to(notebook_dir).parts
            if rel_parts:
                section_key = "/".join(rel_parts) + "/" + section_name
            else:
                section_key = section_name

            if section_key not in sections:
                sections[section_key] = {"files": [], "latest": None}

            sections[section_key]["files"].append(one_file)

        # For each section, determine the latest (most recently modified) file
        for sec_info in sections.values():
            sec_info["files"].sort(key=lambda p: p.stat().st_mtime, reverse=True)
            sec_info["latest"] = sec_info["files"][0]

        if sections:
            notebooks[notebook_name] = {
                "path": notebook_dir,
                "sections": sections,
            }

    return notebooks


def _parse_one_file(filepath: Path) -> list[str]:
    """
    Parse a .one file and extract all text content.

    Returns a list of text strings found in the file.
    """
    texts = []
    try:
        with open(filepath, "rb") as f:
            doc = OneDocment(f)

        props = doc.get_properties()
        for prop in props:
            ptype = prop.get("type", "")
            val = prop.get("val", {})
            if not isinstance(val, dict):
                continue

            # Extract RichEditTextUnicode (the actual text content)
            text = val.get("RichEditTextUnicode", "")
            if text and isinstance(text, str) and text.strip():
                texts.append(text.strip())

    except Exception as e:
        log.warning("Failed to parse %s: %s", filepath, e)

    return texts


def _get_page_titles_from_props(filepath: Path) -> list[str]:
    """Extract page titles from a .one file."""
    titles = []
    try:
        with open(filepath, "rb") as f:
            doc = OneDocment(f)

        props = doc.get_properties()
        for prop in props:
            if prop.get("type") == "jcidTitleNode":
                val = prop.get("val", {})
                if isinstance(val, dict):
                    text = val.get("RichEditTextUnicode", "")
                    if text and text.strip():
                        titles.append(text.strip())
    except Exception as e:
        log.warning("Failed to extract titles from %s: %s", filepath, e)

    return titles


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp = FastMCP("onenote")


@mcp.tool()
async def list_notebooks() -> str:
    """List all locally available OneNote notebooks.

    Shows notebook names and how many sections each one has.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIR}"

    lines = []
    for name, info in sorted(notebooks.items()):
        section_count = len(info["sections"])
        lines.append(f"- {name}  ({section_count} sections)")
    return "\n".join(lines)


@mcp.tool()
async def list_sections(notebook_name: str) -> str:
    """List all sections in a specific notebook.

    Args:
        notebook_name: The name of the notebook (from list_notebooks).
    """
    notebooks = _discover_notebooks()
    if notebook_name not in notebooks:
        # Try case-insensitive match
        for key in notebooks:
            if key.lower() == notebook_name.lower():
                notebook_name = key
                break
        else:
            available = ", ".join(sorted(notebooks.keys()))
            return f"Notebook '{notebook_name}' not found. Available: {available}"

    sections = notebooks[notebook_name]["sections"]
    lines = []
    for sec_name, sec_info in sorted(sections.items()):
        latest = sec_info["latest"]
        size_kb = latest.stat().st_size / 1024
        lines.append(f"- {sec_name}  ({size_kb:.0f} KB)")
    return "\n".join(lines)


@mcp.tool()
async def read_section(notebook_name: str, section_name: str) -> str:
    """Read all text content from a specific section of a notebook.

    Args:
        notebook_name: The name of the notebook.
        section_name: The name of the section (from list_sections).
    """
    notebooks = _discover_notebooks()

    # Case-insensitive notebook match
    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    # Case-insensitive section match
    sec_info = None
    for key, val in nb["sections"].items():
        if key.lower() == section_name.lower():
            sec_info = val
            break
    if sec_info is None:
        available = ", ".join(sorted(nb["sections"].keys()))
        return f"Section '{section_name}' not found. Available: {available}"

    filepath = sec_info["latest"]
    texts = _parse_one_file(filepath)

    if not texts:
        return f"No text content found in section '{section_name}'."

    return "\n\n".join(texts)


@mcp.tool()
async def search_notes(query: str) -> str:
    """Search for text across ALL notebooks and sections.

    Searches through the text content of every section for the given query.
    Returns matching sections with a snippet of the matched text.

    Args:
        query: The text to search for (case-insensitive).
    """
    query_lower = query.lower()
    notebooks = _discover_notebooks()
    results = []

    for nb_name, nb_info in sorted(notebooks.items()):
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            filepath = sec_info["latest"]
            texts = _parse_one_file(filepath)

            for text in texts:
                if query_lower in text.lower():
                    # Build a snippet around the match
                    idx = text.lower().index(query_lower)
                    start = max(0, idx - 80)
                    end = min(len(text), idx + len(query) + 80)
                    snippet = text[start:end].strip()
                    if start > 0:
                        snippet = "..." + snippet
                    if end < len(text):
                        snippet = snippet + "..."

                    results.append(
                        f"[{nb_name} / {sec_name}]\n  {snippet}"
                    )

    if not results:
        return f"No results found for '{query}'."

    header = f"Found {len(results)} match(es) for '{query}':\n\n"
    return header + "\n\n".join(results[:30])  # limit to 30 results


@mcp.tool()
async def list_all_sections() -> str:
    """List ALL sections across ALL notebooks.

    Useful for getting a complete overview of everything in your OneNote.
    """
    notebooks = _discover_notebooks()
    if not notebooks:
        return f"No notebooks found in {ONENOTE_DIR}"

    lines = []
    for nb_name, nb_info in sorted(notebooks.items()):
        lines.append(f"\n## {nb_name}")
        for sec_name, sec_info in sorted(nb_info["sections"].items()):
            latest = sec_info["latest"]
            size_kb = latest.stat().st_size / 1024
            lines.append(f"  - {sec_name}  ({size_kb:.0f} KB)")

    return "\n".join(lines)


@mcp.tool()
async def get_notebook_summary(notebook_name: str) -> str:
    """Get a summary of a notebook: its sections and a preview of each section's content.

    Args:
        notebook_name: The name of the notebook.
    """
    notebooks = _discover_notebooks()

    nb = None
    for key, val in notebooks.items():
        if key.lower() == notebook_name.lower():
            nb = val
            notebook_name = key
            break
    if nb is None:
        available = ", ".join(sorted(notebooks.keys()))
        return f"Notebook '{notebook_name}' not found. Available: {available}"

    lines = [f"# {notebook_name}\n"]

    for sec_name, sec_info in sorted(nb["sections"].items()):
        filepath = sec_info["latest"]
        texts = _parse_one_file(filepath)

        lines.append(f"## {sec_name}")
        if texts:
            # Show first ~200 chars as preview
            preview = " | ".join(texts)
            if len(preview) > 200:
                preview = preview[:200] + "..."
            lines.append(f"  Preview: {preview}")
        else:
            lines.append("  (no text content)")
        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# OneNote COM API helpers (for writing)
# ---------------------------------------------------------------------------

ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _iter_body_oe(root: ET.Element):
    """Yield every <one:OE> element that's part of the page BODY (nested
    inside an <one:Outline>), in document order.

    A page's <one:Title> is ALSO structured as <one:Title><one:OE><one:T>,
    so a plain root.iter("{ns}OE") walk -- which matches the tag anywhere
    in the tree -- picks up the title's OE too, and since Title comes
    before every body Outline in document order, it's whichever OE a
    plain walk finds *first*. find_and_replace_in_page and
    replace_last_block both need "every content block a user would call
    a block", which does not include the title -- this is the one place
    that distinction gets made, so both callers get it for free instead
    of each needing to know to filter Title out themselves.
    """
    for outline in root.iter(f"{{{ONE_NS}}}Outline"):
        yield from outline.iter(f"{{{ONE_NS}}}OE")


def _sanitize_html_for_onenote(html: str) -> str:
    """
    Convert HTML to OneNote-compatible inline HTML.

    OneNote's <one:T> element only supports inline HTML (b, i, span, br, etc.).
    Block-level elements (h1-h6, p, ul, ol, li, div, table, etc.) cause
    UpdatePageContent to silently fail.

    Also escapes ]]> which would break the CDATA wrapper.
    """
    # Escape ]]> so it doesn't break CDATA sections
    html = html.replace("]]>", "]]&gt;")

    # Convert block-level closing tags to <br/>
    html = re.sub(r"</(?:p|div|h[1-6]|li|tr|blockquote|pre|code|section|article|header|footer|nav|aside|details|summary|figure|figcaption|dl|dt|dd)>", "<br/>", html, flags=re.IGNORECASE)

    # Remove block-level opening tags (keep their content)
    html = re.sub(r"<(?:p|div|h[1-6]|li|tr|td|th|blockquote|ul|ol|table|thead|tbody|pre|code|section|article|header|footer|nav|aside|details|summary|figure|figcaption|dl|dt|dd)(?:\s[^>]*)?>", "", html, flags=re.IGNORECASE)

    # Remove remaining closing tags for container elements
    html = re.sub(r"</(?:ul|ol|table|thead|tbody|td|th)>", "", html, flags=re.IGNORECASE)

    # Clean up multiple consecutive <br/> tags
    html = re.sub(r"(<br\s*/?>){3,}", "<br/><br/>", html, flags=re.IGNORECASE)

    # Normalize br tags
    html = re.sub(r"<br\s*/?>", "<br/>", html, flags=re.IGNORECASE)

    # Strip leading/trailing <br/>
    html = re.sub(r"^(<br/>)+", "", html)
    html = re.sub(r"(<br/>)+$", "", html)

    return html.strip()


def _escape_cdata(text: str) -> str:
    """Escape ]]> so it can't prematurely close a CDATA section."""
    return text.replace("]]>", "]]&gt;")


def _run_powershell(script: str) -> tuple[bool, str]:
    """Run a PowerShell script and return (success, output)."""
    try:
        result = subprocess.run(
            ["powershell.exe", "-Command", script],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        if result.returncode != 0:
            return False, result.stderr.strip() or output
        return True, output
    except subprocess.TimeoutExpired:
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        return False, "PowerShell not found (write features require Windows)"


def _com_get_hierarchy(level: int = 3) -> ET.Element | None:
    """
    Get OneNote hierarchy via COM API.
    Levels: 0=Notebooks, 1=SectionGroups, 2=Sections, 3=Sections(full), 4=Pages
    """
    tmpfile = os.path.join(tempfile.gettempdir(), "onenote_hierarchy.xml")
    # Escape backslashes for PowerShell string
    tmpfile_ps = tmpfile.replace("\\", "\\\\")
    script = (
        f'$onenote = New-Object -ComObject OneNote.Application; '
        f'$h = ""; '
        f'$onenote.GetHierarchy("", {level}, [ref]$h); '
        f'$h | Out-File -FilePath "{tmpfile_ps}" -Encoding UTF8; '
        f'Write-Output "OK"'
    )
    ok, msg = _run_powershell(script)
    if not ok:
        log.warning("COM GetHierarchy failed: %s", msg)
        return None
    try:
        with open(tmpfile, "r", encoding="utf-8-sig") as f:
            xml_content = f.read()
        return ET.fromstring(xml_content)
    except Exception as e:
        log.warning("Failed to parse hierarchy XML: %s", e)
        return None
    finally:
        try:
            os.remove(tmpfile)
        except OSError:
            pass


def _com_get_page_content(page_id: str) -> str | None:
    """Fetch the raw page content XML for a live page via the COM API."""
    tmpfile = os.path.join(tempfile.gettempdir(), "onenote_page_content.xml")
    tmpfile_ps = tmpfile.replace("\\", "\\\\")
    page_id_esc = page_id.replace("'", "''")
    script = (
        f'$onenote = New-Object -ComObject OneNote.Application; '
        f'$p = ""; '
        f"$onenote.GetPageContent('{page_id_esc}', [ref]$p, 0); "
        f'$p | Out-File -FilePath "{tmpfile_ps}" -Encoding UTF8; '
        f'Write-Output "OK"'
    )
    ok, msg = _run_powershell(script)
    if not ok:
        log.warning("COM GetPageContent failed for %s: %s", page_id, msg)
        return None
    try:
        with open(tmpfile, "r", encoding="utf-8-sig") as f:
            return f.read()
    except Exception as e:
        log.warning("Failed to read page content file: %s", e)
        return None
    finally:
        try:
            os.remove(tmpfile)
        except OSError:
            pass


def _extract_text_from_page_xml(xml_content: str) -> str:
    """Extract plain text from a OneNote page-content XML document.

    Text lives in <one:T> elements, often as inline HTML (spans, bold, etc.),
    so tags are stripped and HTML entities unescaped after extraction.

    OneNote also runs OCR on embedded images (e.g. a stat block screenshot,
    a scanned PDF page) and stores the recognized text in a nested
    <one:OCRData><one:OCRText> element right next to the <one:Image>. That
    text is picked up here too -- and clearly labeled, since OCR output can
    contain recognition errors -- so image content isn't silently invisible
    to the text-only reading tools.

    Walking root.iter() with no tag filter visits every element in document
    order, so <one:T> and <one:OCRText> content comes out roughly in the
    order it appears on the page instead of images being lumped separately.
    """
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        log.warning("Failed to parse page content XML: %s", e)
        return ""

    t_tag = f"{{{ONE_NS}}}T"
    ocr_tag = f"{{{ONE_NS}}}OCRText"

    lines = []
    for el in root.iter():
        if el.tag == t_tag:
            raw = el.text or ""
            if not raw.strip():
                continue
            plain = re.sub(r"<[^>]+>", "", raw)
            plain = html.unescape(plain).strip()
            if plain:
                lines.append(plain)
        elif el.tag == ocr_tag:
            raw = el.text or ""
            plain = html.unescape(raw).strip() if raw else ""
            if plain:
                lines.append(f"[OCR text from an image, may contain recognition errors]\n{plain}")
    return "\n".join(lines)


def _com_find_section_id(
    notebook_name: str, section_name: str, include_recycled: bool = False
) -> str | None:
    """Find a section ID by notebook and section name (case-insensitive).

    Searches recursively through nested section groups. Recycle-bin items are
    skipped by default; pass include_recycled=True to also find (and thereby
    be able to restore) a deleted section.
    """
    root = _com_get_hierarchy(3)
    if root is None:
        return None

    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        for sec in nb.iter(f"{{{ONE_NS}}}Section"):
            if not include_recycled and sec.get("isInRecycleBin") == "true":
                continue
            if sec.get("name", "").lower() == section_name.lower():
                return sec.get("ID")
    return None


def _com_find_section_group_id(
    notebook_name: str, group_name: str, include_recycled: bool = False
) -> str | None:
    """Find a section group (folder) ID by notebook and group name (case-insensitive).

    Searches recursively through nested section groups.
    """
    root = _com_get_hierarchy(3)
    if root is None:
        return None

    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        for grp in nb.iter(f"{{{ONE_NS}}}SectionGroup"):
            if not include_recycled and grp.get("isInRecycleBin") == "true":
                continue
            if grp.get("name", "").lower() == group_name.lower():
                return grp.get("ID")
    return None


def _com_find_section_parent(notebook_name: str, section_id: str) -> tuple[str, str] | None:
    """Find a section's current direct parent, as (parent_id, parent_tag)
    where parent_tag is "Notebook" or "SectionGroup". Used to remember
    where a section was before moving or deleting it, so that move can
    later be undone.
    """
    root = _com_get_hierarchy(3)
    if root is None:
        return None

    def walk(container: ET.Element, container_id: str, container_tag: str) -> tuple[str, str] | None:
        for child in container:
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "Section" and child.get("ID") == section_id:
                return (container_id, container_tag)
            if tag == "SectionGroup":
                found = walk(child, child.get("ID", ""), "SectionGroup")
                if found:
                    return found
        return None

    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        return walk(nb, nb.get("ID", ""), "Notebook")
    return None


def _format_hierarchy_tree(node: ET.Element, depth: int = 0) -> list[str]:
    """Recursively render a Notebook/SectionGroup element's children as an
    indented tree of section-group folders and sections, skipping recycle-bin
    items."""
    lines = []
    indent = "  " * depth
    for child in node:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "SectionGroup":
            if child.get("isInRecycleBin") == "true":
                continue
            lines.append(f"{indent}- {child.get('name', '?')}/  (group)")
            lines.extend(_format_hierarchy_tree(child, depth + 1))
        elif tag == "Section":
            if child.get("isInRecycleBin") == "true":
                continue
            locked = " (locked)" if child.get("locked") == "true" else ""
            lines.append(f"{indent}- {child.get('name', '?')}{locked}")
    return lines


def _com_find_notebook_id(notebook_name: str) -> str | None:
    """Find a notebook's ID by name (case-insensitive)."""
    # HierarchyScope 0 ("Notebooks") returns an empty result on some OneNote
    # versions; level 3 reliably includes Notebook elements with an ID.
    root = _com_get_hierarchy(3)
    if root is None:
        return None
    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() == notebook_name.lower():
            return nb.get("ID")
    return None


def _run_powershell_file(script: str) -> tuple[bool, str]:
    """Write a PowerShell script to a temp file and execute it."""
    ps_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_cmd.ps1")
    try:
        with open(ps_file, "w", encoding="utf-8") as f:
            f.write(script)
        log.debug("Running PowerShell script (%d chars): %s", len(script), ps_file)
        result = subprocess.run(
            ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True, text=True, timeout=30,
        )
        output = result.stdout.strip()
        stderr = result.stderr.strip()
        log.debug("PowerShell exit=%d stdout=%s stderr=%s",
                  result.returncode, output[:500] if output else "(empty)",
                  stderr[:500] if stderr else "(empty)")
        if result.returncode != 0:
            return False, stderr or output
        return True, output
    except subprocess.TimeoutExpired:
        log.error("PowerShell timed out")
        return False, "PowerShell command timed out"
    except FileNotFoundError:
        log.error("PowerShell not found")
        return False, "PowerShell not found (write features require Windows)"
    finally:
        try:
            os.remove(ps_file)
        except OSError:
            pass


def _com_create_section(notebook_id: str, section_name: str) -> tuple[bool, str]:
    """Create a new section in a notebook using the OneNote COM API.

    Uses OpenHierarchy with newObjectType=3 (cftSection). When bstrPath is
    passed relative to a parent object ID, it must be a filename ending in
    ".one" -- passing just the bare section name fails with an HRESULT error.
    """
    section_file = f"{section_name}.one"
    section_file_esc = section_file.replace("'", "''")
    notebook_id_esc = notebook_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$newId = ""
try {{
    $onenote.OpenHierarchy('{section_file_esc}', '{notebook_id_esc}', [ref]$newId, 3)
}} catch {{
    Write-Error "OpenHierarchy failed: $_"
    exit 1
}}
Write-Output $newId
"""
    ok, output = _run_powershell_file(script)
    log.info("create_section: notebook=%s name=%r ok=%s output=%r", notebook_id, section_name, ok, output)
    if ok and output:
        return True, output
    return False, output or "Unknown error creating section"


def _com_create_section_group(notebook_id: str, group_name: str) -> tuple[bool, str]:
    """Create a new section group (folder) in a notebook using OpenHierarchy
    with newObjectType=2 (cftFolder). Unlike sections, group paths take no
    file extension.
    """
    group_name_esc = group_name.replace("'", "''")
    notebook_id_esc = notebook_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$newId = ""
try {{
    $onenote.OpenHierarchy('{group_name_esc}', '{notebook_id_esc}', [ref]$newId, 2)
}} catch {{
    Write-Error "OpenHierarchy failed: $_"
    exit 1
}}
Write-Output $newId
"""
    ok, output = _run_powershell_file(script)
    log.info("create_section_group: notebook=%s name=%r ok=%s output=%r", notebook_id, group_name, ok, output)
    if ok and output:
        return True, output
    return False, output or "Unknown error creating section group"


def _com_move_section(section_id: str, destination_id: str, destination_tag: str) -> tuple[bool, str]:
    """Move an existing section (by ID) under a different parent (notebook or
    section group, by ID) using UpdateHierarchy. Wrapping the existing
    section's ID as a child of the target parent reparents it instead of
    creating a duplicate. Also works to move a section out of the recycle
    bin (restore), since the recycle bin is itself a section group.

    destination_tag must be "Notebook" or "SectionGroup" to match what
    destination_id actually refers to -- using the wrong wrapper element
    fails with an HRESULT error.
    """
    section_id_esc = section_id.replace("'", "''")
    destination_id_esc = destination_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = "<one:{destination_tag} xmlns:one=`"http://schemas.microsoft.com/office/onenote/2013/onenote`" ID=`"{destination_id_esc}`"><one:Section ID=`"{section_id_esc}`" /></one:{destination_tag}>"
try {{
    $onenote.UpdateHierarchy($xml)
}} catch {{
    Write-Error "UpdateHierarchy failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    ok, output = _run_powershell_file(script)
    log.info("move_section: section=%s dest=%s ok=%s output=%r", section_id, destination_id, ok, output)
    if ok:
        return True, "Section moved successfully."
    return False, output or "Unknown error moving section"


def _com_rename_section(section_id: str, new_name: str) -> tuple[bool, str]:
    """Rename an existing section (by ID) using UpdateHierarchy."""
    section_id_esc = section_id.replace("'", "''")
    new_name_esc = new_name.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
$xml = "<one:Section xmlns:one=`"http://schemas.microsoft.com/office/onenote/2013/onenote`" ID=`"{section_id_esc}`" name=`"{new_name_esc}`" />"
try {{
    $onenote.UpdateHierarchy($xml)
}} catch {{
    Write-Error "UpdateHierarchy failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    ok, output = _run_powershell_file(script)
    log.info("rename_section: section=%s new_name=%r ok=%s output=%r", section_id, new_name, ok, output)
    if ok:
        return True, "Section renamed successfully."
    return False, output or "Unknown error renaming section"


def _com_delete_hierarchy(object_id: str) -> tuple[bool, str]:
    """Delete a hierarchy object (section or page) by ID using DeleteHierarchy.
    Moves to the OneNote recycle bin (does not pass deletePermanently)."""
    object_id_esc = object_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
try {{
    $onenote.DeleteHierarchy('{object_id_esc}')
}} catch {{
    Write-Error "DeleteHierarchy failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    ok, output = _run_powershell_file(script)
    log.info("delete_hierarchy: object=%s ok=%s output=%r", object_id, ok, output)
    if ok:
        return True, "Deleted (moved to OneNote recycle bin)."
    return False, output or "Unknown error deleting"


def _com_create_page(section_id: str, title: str, body_html: str) -> tuple[bool, str]:
    """Create a new page in a section using the OneNote COM API."""
    log.info("create_page: title=%r, body_len=%d, section=%s", title, len(body_html), section_id)
    log.debug("create_page: raw body=%r", body_html[:500])
    # Sanitize HTML to OneNote-compatible inline format
    body_html = _sanitize_html_for_onenote(body_html)
    log.debug("create_page: sanitized body=%r", body_html[:500])

    # Write title and body to temp files to avoid all escaping issues
    title_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_title.txt")
    body_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_body.txt")
    with open(title_file, "w", encoding="utf-8") as f:
        f.write(_escape_cdata(title))
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(body_html)

    section_id_esc = section_id.replace("'", "''")

    script = f"""
$titleContent = Get-Content -Path '{title_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
$bodyContent = Get-Content -Path '{body_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($titleContent) {{ $titleContent = $titleContent.Trim() }}
if ($bodyContent) {{ $bodyContent = $bodyContent.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageId = ""
$onenote.CreateNewPage('{section_id_esc}', [ref]$pageId, 0)

# Get the new page's XML
$pageXml = ""
$onenote.GetPageContent($pageId, [ref]$pageXml, 0)
$xml = [xml]$pageXml

# Set title
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$titleNode = $xml.SelectSingleNode("//one:Title/one:OE/one:T", $nsMgr)
if ($titleNode) {{
    $titleNode.InnerXml = "<![CDATA[" + $titleContent + "]]>"
}}

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Title UpdatePageContent failed: $_"
    exit 1
}}

# Re-fetch to add body
$pageXml2 = ""
$onenote.GetPageContent($pageId, [ref]$pageXml2, 0)
$xml2 = [xml]$pageXml2

# Add body outline
$outline = $xml2.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml2.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml2.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml2.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml2.CreateCDataSection($bodyContent)
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml2.DocumentElement.AppendChild($outline) | Out-Null

try {{
    $onenote.UpdatePageContent($xml2.OuterXml)
}} catch {{
    Write-Error "Body UpdatePageContent failed: $_"
    exit 1
}}
Write-Output $pageId
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("create_page result: ok=%s output=%r", ok, output[:200] if output else "(empty)")
        if ok and output:
            return True, f"Page '{title}' created successfully (ID: {output})"
        return False, f"Failed to create page: {output}"
    finally:
        for f in (title_file, body_file):
            try:
                os.remove(f)
            except OSError:
                pass


def _com_get_page_title(page_id: str) -> str | None:
    """Fetch a live page's current title text (for capturing undo data
    before renaming it)."""
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return None
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        return None
    title_t = root.find(f"{{{ONE_NS}}}Title/{{{ONE_NS}}}OE/{{{ONE_NS}}}T")
    if title_t is None or not title_t.text:
        return None
    return html.unescape(re.sub(r"<[^>]+>", "", title_t.text)).strip()


def _com_rename_page(page_id: str, new_title: str) -> tuple[bool, str]:
    """Rename an existing page's title, leaving its body content untouched."""
    title_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_title.txt")
    with open(title_file, "w", encoding="utf-8") as f:
        f.write(_escape_cdata(new_title))

    page_id_esc = page_id.replace("'", "''")
    script = f"""
$titleContent = Get-Content -Path '{title_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($titleContent) {{ $titleContent = $titleContent.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml

$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$titleNode = $xml.SelectSingleNode("//one:Title/one:OE/one:T", $nsMgr)
if ($titleNode) {{
    $titleNode.InnerXml = "<![CDATA[" + $titleContent + "]]>"
}}

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Title UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("rename_page: page=%s new_title=%r ok=%s output=%r", page_id, new_title, ok, output)
        if ok:
            return True, "Page renamed successfully."
        return False, output or "Unknown error renaming page"
    finally:
        try:
            os.remove(title_file)
        except OSError:
            pass


def _com_insert_image_bytes(page_id: str, raw_bytes: bytes, image_format: str) -> tuple[bool, str]:
    """Insert an image (given as raw bytes) as a new content block on an
    existing page, using the OneNote COM API.

    image_format is the file format the bytes are in (e.g. "png", "jpeg",
    "gif") -- OneNote needs this to know how to decode/render it.

    Without an explicit <one:Size>, OneNote auto-scales inserted images down
    to a small default (observed: ~250x250), so the image's real pixel
    dimensions are read via Pillow and set with isSetByUser="true" to stop
    OneNote from touching the size.
    """
    image_base64 = base64.b64encode(raw_bytes).decode("ascii")
    log.info("insert_image: page=%s, format=%s, bytes=%d", page_id, image_format, len(raw_bytes))

    try:
        with PILImage.open(io.BytesIO(raw_bytes)) as img:
            width_px, height_px = img.size
    except Exception as e:
        log.warning("insert_image: could not read image dimensions: %s", e)
        width_px, height_px = None, None

    # Base64 data can be large; write to a temp file to avoid PowerShell
    # command-line/argument length limits, same as title/body already do.
    b64_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_image_b64.txt")
    with open(b64_file, "w", encoding="ascii") as f:
        f.write(image_base64)

    page_id_esc = page_id.replace("'", "''")
    format_esc = image_format.replace("'", "''")

    if width_px and height_px:
        size_line = (
            f'$size = $xml.CreateElement("one", "Size", "http://schemas.microsoft.com/office/onenote/2013/onenote"); '
            f'$size.SetAttribute("width", "{width_px}.0"); '
            f'$size.SetAttribute("height", "{height_px}.0"); '
            f'$size.SetAttribute("isSetByUser", "true"); '
            f'$image.AppendChild($size) | Out-Null'
        )
    else:
        size_line = "# image dimensions unknown; letting OneNote pick a size"

    script = f"""
$imageB64 = Get-Content -Path '{b64_file.replace(chr(39), chr(39)+chr(39))}' -Raw
if ($imageB64) {{ $imageB64 = $imageB64.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml

$outline = $xml.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$image = $xml.CreateElement("one", "Image", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$image.SetAttribute("format", "{format_esc}")
{size_line}
$data = $xml.CreateElement("one", "Data", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$data.InnerText = $imageB64
$image.AppendChild($data) | Out-Null
$oe.AppendChild($image) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml.DocumentElement.AppendChild($outline) | Out-Null

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Insert image UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("insert_image result: ok=%s output=%r", ok, output[:200] if output else "(empty)")
        if ok:
            return True, "Image inserted successfully."
        return False, f"Failed to insert image: {output}"
    finally:
        try:
            os.remove(b64_file)
        except OSError:
            pass


def _com_insert_image_from_file(page_id: str, file_path: str) -> tuple[bool, str]:
    """Insert an image by reading it from a local file. The OneNote COM API
    embeds the file's bytes directly -- the caller only ever needs to pass a
    path, never the image data itself.

    If file_path doesn't exist as given (e.g. it's just a bare filename, not
    a full path), it's also looked up inside IMAGE_DROP_DIR -- so a user can
    save "Background.png" wherever that is (Downloads by default) and refer
    to it by name alone.
    """
    path = Path(file_path)
    if not path.is_file():
        fallback = IMAGE_DROP_DIR / file_path
        if fallback.is_file():
            path = fallback
        else:
            return False, (
                f"File not found: {file_path} "
                f"(also checked {IMAGE_DROP_DIR})"
            )

    try:
        raw_bytes = path.read_bytes()
    except OSError as e:
        return False, f"Could not read file: {e}"

    try:
        with PILImage.open(io.BytesIO(raw_bytes)) as img:
            image_format = (img.format or "").lower()
    except Exception:
        # Fall back to the file extension if Pillow can't identify it.
        image_format = path.suffix.lstrip(".").lower()

    # OneNote's <one:Image format="..."> attribute only accepts a fixed
    # vocabulary (auto, bmp, gif, jpg, png, emf, wdp, ...) -- Pillow's own
    # vocabulary calls the same format "jpeg", which OneNote's schema
    # rejects outright, so insert_image_from_file failed for every .jpg
    # file even though PNG worked fine.
    if image_format == "jpeg":
        image_format = "jpg"

    if not image_format:
        return False, f"Could not determine image format for: {file_path}"

    return _com_insert_image_bytes(page_id, raw_bytes, image_format)


def _com_find_and_replace_in_page(
    page_id: str, find_text: str, replace_text: str, replace_all: bool = False
) -> tuple[bool, str, int]:
    """Find and replace text within a page's BODY -- never its title, see
    _iter_body_oe.

    Matches against each <one:OE> content block's whole plain text (its
    direct <one:T> runs joined together first), so a match split across
    sibling runs by inline formatting (e.g. text interrupted by <b>) is
    still found. Replacing flattens that OE's runs into a single plain
    text node, so any inline formatting *within that specific matched
    block* is lost. Exact substring match only -- no fuzzy matching, so a
    non-match is always reported honestly rather than guessed at.

    Returns (ok, message, replacement_count).
    """
    if not find_text:
        return False, "find_text must not be empty.", 0

    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return False, "Could not read page content.", 0

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return False, f"Could not parse page XML: {e}", 0

    t_tag = f"{{{ONE_NS}}}T"

    targets = []
    total = 0
    oe_index = -1
    for oe in _iter_body_oe(root):
        oe_index += 1
        t_children = [c for c in list(oe) if c.tag == t_tag]
        if not t_children:
            continue

        raw_joined = "".join(t.text or "" for t in t_children)
        joined = html.unescape(re.sub(r"<[^>]+>", "", raw_joined))
        if find_text not in joined:
            continue

        if replace_all:
            n = joined.count(find_text)
            new_joined = joined.replace(find_text, replace_text)
        else:
            new_joined = joined.replace(find_text, replace_text, 1)
            n = 1

        targets.append({
            "index": oe_index,
            "text": _escape_cdata(new_joined),
            "expected_raw": raw_joined,
        })
        total += n
        if not replace_all:
            break

    if not targets:
        return True, "No match found; nothing was changed.", 0

    targets_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_replace_targets.json")
    with open(targets_file, "w", encoding="utf-8") as f:
        json.dump(targets, f)

    page_id_esc = page_id.replace("'", "''")
    script = f"""
$targetsJson = Get-Content -Path '{targets_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
$targets = $targetsJson | ConvertFrom-Json

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeNodes = $xml.SelectNodes("//one:Outline//one:OE", $nsMgr)

# Verify every target still has the exact content it had when Python
# analyzed the page, before changing anything. GetPageContent is called
# twice (once in Python, once here) as two separate COM round-trips; if the
# page changed in between (a concurrent edit), applying the old plan could
# silently hit the wrong block. Abort entirely rather than risk that.
foreach ($target in $targets) {{
    $oe = $oeNodes[$target.index]
    if ($oe -eq $null) {{
        Write-Error "Page structure changed since it was read (block no longer exists) -- aborting, nothing was modified."
        exit 1
    }}
    $tNodes = @($oe.ChildNodes | Where-Object {{ $_.LocalName -eq "T" }})
    $actualRaw = -join ($tNodes | ForEach-Object {{ $_.InnerText }})
    if ($actualRaw -ne $target.expected_raw) {{
        Write-Error "Page content changed since it was read -- aborting, nothing was modified. Please retry."
        exit 1
    }}
}}

foreach ($target in $targets) {{
    $oe = $oeNodes[$target.index]
    $tNodes = @($oe.ChildNodes | Where-Object {{ $_.LocalName -eq "T" }})
    foreach ($tn in $tNodes) {{ $oe.RemoveChild($tn) | Out-Null }}
    $newT = $xml.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
    $cdata = $xml.CreateCDataSection($target.text)
    $newT.AppendChild($cdata) | Out-Null
    $oe.AppendChild($newT) | Out-Null
}}

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Find/replace UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("find_and_replace_in_page: page=%s ok=%s targets=%d output=%r",
                  page_id, ok, len(targets), output[:200] if output else "(empty)")
        if ok:
            return True, "Replacement successful.", total
        return False, f"Failed to update page: {output}", 0
    finally:
        try:
            os.remove(targets_file)
        except OSError:
            pass


def _com_replace_block_by_id(
    page_id: str, block_id: str, new_content: str, expected_raw: str | None = None
) -> tuple[bool, str, str]:
    """Replace one specific body content block, addressed by its stable
    objectID rather than by its position in the page.

    This is the core mutation _com_replace_last_block delegates to (after
    resolving what "the last block" currently means) -- and what undo/
    redo of a previous replace_last_block call use directly, since they
    already know exactly which block was changed and shouldn't
    re-resolve "the last block" fresh: the page may have grown new
    blocks since (another append_to_page or replace_last_block call in
    between), which would otherwise make undo/redo silently act on the
    wrong block.

    expected_raw, when given, is verified against the block's current raw
    content right before writing (same optimistic-concurrency check
    find_and_replace_in_page and this function's caller have always done)
    -- aborts rather than risk overwriting content that changed for an
    unrelated reason since expected_raw was captured.

    Returns (ok, message, previous_plain_text).
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return False, "Could not read page content.", ""

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return False, f"Could not parse page XML: {e}", ""

    t_tag = f"{{{ONE_NS}}}T"
    target_oe = None
    for oe in _iter_body_oe(root):
        if oe.get("objectID") == block_id:
            target_oe = oe
            break
    if target_oe is None:
        return False, f"Could not find a block with ID {block_id} -- it may have been deleted, or the page changed.", ""

    raw_old_text = "".join(c.text or "" for c in list(target_oe) if c.tag == t_tag)
    old_text = html.unescape(re.sub(r"<[^>]+>", "", raw_old_text))
    verify_against = raw_old_text if expected_raw is None else expected_raw

    body_html = _escape_cdata(_sanitize_html_for_onenote(new_content))
    body_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_replace_block_body.txt")
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(body_html)
    expected_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_replace_block_expected.txt")
    # newline="" disables Python's universal-newline translation on write, so
    # the file holds the exact original bytes -- needed since this is
    # compared byte-for-byte against the live XML's InnerText below.
    with open(expected_file, "w", encoding="utf-8", newline="") as f:
        f.write(verify_against)

    page_id_esc = page_id.replace("'", "''")
    block_id_esc = block_id.replace("'", "''")
    script = f"""
$bodyContent = Get-Content -Path '{body_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($bodyContent) {{ $bodyContent = $bodyContent.Trim() }}
$expectedRaw = Get-Content -Path '{expected_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($expectedRaw -eq $null) {{ $expectedRaw = "" }}

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml
$nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsMgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$targetOe = $xml.SelectSingleNode("//one:Outline//one:OE[@objectID='{block_id_esc}']", $nsMgr)
if ($targetOe -eq $null) {{
    Write-Error "Block {block_id_esc} no longer exists -- aborting, nothing was modified."
    exit 1
}}
$tNodes = @($targetOe.ChildNodes | Where-Object {{ $_.LocalName -eq "T" }})
$actualRaw = -join ($tNodes | ForEach-Object {{ $_.InnerText }})
if ($actualRaw -ne $expectedRaw) {{
    Write-Error "Block content changed since it was read -- aborting, nothing was modified. Please retry."
    exit 1
}}
foreach ($tn in $tNodes) {{ $targetOe.RemoveChild($tn) | Out-Null }}
$newT = $xml.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml.CreateCDataSection($bodyContent)
$newT.AppendChild($cdata) | Out-Null
$targetOe.AppendChild($newT) | Out-Null

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Replace-block-by-id UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("replace_block_by_id: page=%s block=%s ok=%s output=%r",
                  page_id, block_id, ok, output[:200] if output else "(empty)")
        if ok:
            return True, "Block replaced successfully.", old_text
        return False, f"Failed to replace block: {output}", ""
    finally:
        for f in (body_file, expected_file):
            try:
                os.remove(f)
            except OSError:
                pass


def _com_count_body_text_blocks(page_id: str) -> int | None:
    """Count how many body content blocks (text-bearing <one:OE> elements
    -- same definition _com_replace_last_block uses for "a block") a page
    currently has. Returns None if the page couldn't be read.

    Used by the replace_last_block tool's confirmation gate to detect the
    single-block case -- where "replace the last block" actually means
    "replace the entire page body" -- before making any change.
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return None
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        return None
    t_tag = f"{{{ONE_NS}}}T"
    return sum(1 for oe in _iter_body_oe(root) if any(c.tag == t_tag for c in list(oe)))


def _com_replace_last_block(page_id: str, new_content: str) -> tuple[bool, str, str, str]:
    """Replace the last <one:OE> content block in a page's BODY with
    new_content, regardless of what it previously contained. Never
    touches the page's title, even though the title is structurally also
    an <one:OE> -- see _iter_body_oe.

    "Block" here means one <one:OE> element, OneNote's own unit of content.
    A page built from a single create_page call (the common case) has
    exactly one such block holding the *entire* body -- so on a page like
    that, this replaces the whole page body, not just a trailing sentence
    or paragraph. The replace_last_block TOOL wrapper is responsible for
    gating that single-block case behind confirm=true -- this function
    always performs the replacement it's asked for.

    Resolves "the last block" once here (by position, since that's the
    only way to define "last"), then delegates the actual write to
    _com_replace_block_by_id using that block's stable objectID -- so a
    later undo/redo of this specific call can target the exact same
    block directly, instead of re-resolving "last" fresh (which could
    mean a different block after further edits).

    Returns (ok, message, previous_plain_text, block_id).
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return False, "Could not read page content.", "", ""

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return False, f"Could not parse page XML: {e}", "", ""

    t_tag = f"{{{ONE_NS}}}T"
    all_oe = list(_iter_body_oe(root))

    # "Last block" means the last OE that actually has text in it -- a
    # trailing image-only OE (e.g. from insert_image_from_file) shouldn't
    # be mistaken for the last-written text block.
    last_index = None
    for i in range(len(all_oe) - 1, -1, -1):
        if any(c.tag == t_tag for c in list(all_oe[i])):
            last_index = i
            break
    if last_index is None:
        return False, "This page has no content blocks to replace.", "", ""

    last_oe = all_oe[last_index]
    raw_old_text = "".join(c.text or "" for c in list(last_oe) if c.tag == t_tag)
    is_only_block = len(all_oe) == 1
    block_id = last_oe.get("objectID", "")

    if not block_id:
        # Every block a live GetPageContent call returns should carry a
        # real objectID -- OneNote assigns one the moment content becomes
        # part of the live page, not just on next explicit save. Bail
        # out honestly rather than fabricate a way to target it.
        return False, "This block has no stable ID (unexpected OneNote state) -- cannot safely target it.", "", ""

    ok, msg, old_text = _com_replace_block_by_id(page_id, block_id, new_content, expected_raw=raw_old_text)
    if ok and is_only_block:
        msg += (
            " Note: this page had only one content block, so this "
            "replaced the entire page body, not just a trailing part of it."
        )
    return ok, msg, old_text, block_id


def _com_delete_page_content_object(page_id: str, object_id: str) -> tuple[bool, str]:
    """Delete one page-level object (an Outline, Image, or Ink block) by its
    objectID, via the dedicated DeletePageContent method.

    UpdatePageContent cannot be used for removal -- confirmed by testing,
    omitting an element from the XML passed to it does not delete it, it's
    only additive/modifying. DeletePageContent is the correct API for
    actually removing a page-level object.
    """
    page_id_esc = page_id.replace("'", "''")
    object_id_esc = object_id.replace("'", "''")
    script = f"""
$onenote = New-Object -ComObject OneNote.Application
try {{
    $onenote.DeletePageContent('{page_id_esc}', '{object_id_esc}')
}} catch {{
    Write-Error "DeletePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    ok, output = _run_powershell_file(script)
    log.info("delete_page_content_object: page=%s object=%s ok=%s output=%r",
              page_id, object_id, ok, output[:200] if output else "(empty)")
    if ok:
        return True, "Block removed successfully."
    return False, f"Failed to remove block: {output}"


def _com_remove_last_block(page_id: str) -> tuple[bool, str]:
    """Remove the last top-level <one:Outline> content block on a page
    entirely. Fallback for when a specific objectID isn't known -- prefer
    _com_delete_page_content_object with an explicit ID when one is
    available (e.g. from append_to_page's return value), since "whatever is
    currently last" can be the wrong block if something else was added to
    the page after the block this call is meant to target.
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return False, "Could not read page content."

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return False, f"Could not parse page XML: {e}"

    outlines = list(root.iter(f"{{{ONE_NS}}}Outline"))
    if not outlines:
        return False, "This page has no content blocks to remove."

    last_outline_id = outlines[-1].get("objectID")
    if not last_outline_id:
        return False, "Could not determine the last block's object ID."

    return _com_delete_page_content_object(page_id, last_outline_id)


def _com_append_to_page(page_id: str, body_html: str) -> tuple[bool, str, str]:
    """Append content to an existing page using the OneNote COM API.

    Returns (ok, message, new_outline_id). new_outline_id is the objectID of
    the freshly-created Outline block, captured immediately via a follow-up
    GetPageContent call right after the append succeeds -- this lets undo
    later target that *specific* block precisely, instead of assuming
    whatever is "currently last" at undo time is still the right one (which
    breaks if something else got added to the page in between).
    """
    log.info("append_to_page: page=%s, body_len=%d", page_id, len(body_html))
    log.debug("append_to_page: raw body=%r", body_html[:500])
    body_html = _sanitize_html_for_onenote(body_html)
    log.debug("append_to_page: sanitized body=%r", body_html[:500])

    # Write body to temp file to avoid escaping issues
    body_file = os.path.join(tempfile.gettempdir(), "onenote_mcp_body.txt")
    with open(body_file, "w", encoding="utf-8") as f:
        f.write(body_html)

    page_id_esc = page_id.replace("'", "''")

    script = f"""
$bodyContent = Get-Content -Path '{body_file.replace(chr(39), chr(39)+chr(39))}' -Raw -Encoding UTF8
if ($bodyContent) {{ $bodyContent = $bodyContent.Trim() }}

$onenote = New-Object -ComObject OneNote.Application
$pageXml = ""
$onenote.GetPageContent('{page_id_esc}', [ref]$pageXml, 0)
$xml = [xml]$pageXml

$outline = $xml.CreateElement("one", "Outline", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oeChildren = $xml.CreateElement("one", "OEChildren", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$oe = $xml.CreateElement("one", "OE", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$t = $xml.CreateElement("one", "T", "http://schemas.microsoft.com/office/onenote/2013/onenote")
$cdata = $xml.CreateCDataSection($bodyContent)
$t.AppendChild($cdata) | Out-Null
$oe.AppendChild($t) | Out-Null
$oeChildren.AppendChild($oe) | Out-Null
$outline.AppendChild($oeChildren) | Out-Null
$xml.DocumentElement.AppendChild($outline) | Out-Null

try {{
    $onenote.UpdatePageContent($xml.OuterXml)
}} catch {{
    Write-Error "Append UpdatePageContent failed: $_"
    exit 1
}}
Write-Output "OK"
"""
    try:
        ok, output = _run_powershell_file(script)
        log.info("append_to_page result: ok=%s output=%r", ok, output[:200] if output else "(empty)")
        if not ok:
            return False, f"Failed to append content: {output}", ""
    finally:
        try:
            os.remove(body_file)
        except OSError:
            pass

    new_outline_id = ""
    fresh_xml = _com_get_page_content(page_id)
    if fresh_xml:
        try:
            fresh_root = ET.fromstring(fresh_xml)
            outlines = list(fresh_root.iter(f"{{{ONE_NS}}}Outline"))
            if outlines:
                new_outline_id = outlines[-1].get("objectID", "")
        except ET.ParseError as e:
            log.warning("append_to_page: could not re-parse page to capture new block ID: %s", e)

    return True, "Content appended successfully.", new_outline_id


def _splice_paragraph_after(
    raw_html: str, anchor_text: str, new_paragraph_html: str
) -> tuple[str | None, int]:
    """Given one OE's raw (already OneNote-sanitized, <br/>-delimited)
    content, split it into paragraph-like segments on runs of <br/> tags,
    find the ONE segment whose plain text contains anchor_text, and
    return new raw content with new_paragraph_html spliced in as a fresh
    paragraph immediately after it -- preserving every other segment's
    original markup untouched (only the two new boundaries around the
    inserted paragraph are added; existing separators elsewhere are kept
    as-is).

    This is what makes insert_block_after actually insert IN PLACE within
    a page's reading flow, rather than only after whichever <one:Outline>
    the anchor happens to sit in -- a page written by a single create_page
    call has its ENTIRE body as one single OE/Outline with multiple
    <br/>-separated paragraphs inside it, so anchoring at the Outline
    level would always land at the end of the page, indistinguishable
    from append_to_page, regardless of where in the text anchor_text
    actually was.

    Returns (new_raw_html, match_count). new_raw_html is None when
    match_count != 1 (0 = anchor not found in this block at all, >1 =
    anchor_text appears in more than one paragraph within this same
    block -- both reported honestly by the caller rather than guessed at).
    """
    parts = re.split(r"((?:<br/>)+)", raw_html)
    segments = parts[0::2]
    seps = parts[1::2]
    n = len(segments)

    match_indices = [
        i for i, seg in enumerate(segments)
        if anchor_text in html.unescape(re.sub(r"<[^>]+>", "", seg))
    ]
    if len(match_indices) != 1:
        return None, len(match_indices)

    idx = match_indices[0]
    default_sep = seps[0] if seps else "<br/><br/>"

    rebuilt = []
    for i in range(n):
        rebuilt.append(segments[i])
        if i == idx:
            rebuilt.append(default_sep)
            rebuilt.append(new_paragraph_html)
            if i < n - 1:
                rebuilt.append(default_sep)
        elif i < n - 1:
            rebuilt.append(seps[i])
    return "".join(rebuilt), 1


def _com_insert_block_after(
    page_id: str, anchor_text: str, new_content: str
) -> tuple[bool, str, str, str, str]:
    """Insert new_content as a new paragraph immediately after the
    existing BODY text containing anchor_text -- for putting something in
    a specific place partway through a page, not at the end
    (append_to_page) and not by replacing something already there
    (replace_last_block/find_and_replace_in_page). Never matches inside
    the page's title -- see _iter_body_oe.

    Inserts WITHIN the matched block's own content (splitting it into
    <br/>-delimited paragraphs and splicing the new one in after the
    matching paragraph), not as a new sibling <one:Outline> -- a page
    written by a single create_page call has its whole body as one OE, so
    anchoring at the Outline/block level would always land at the very
    end of the page regardless of where anchor_text actually is. This
    way "insert after X" means what it says on every page shape.

    anchor_text must match exactly (case-sensitive, no fuzzy matching),
    same contract as find_and_replace_in_page -- if it matches more than
    one paragraph (whether in the same block or different ones), or none
    at all, that's reported honestly rather than guessed at.

    Delegates the actual write to _com_replace_block_by_id, targeting the
    matched block's stable objectID with its recomputed full content --
    reusing the same temp-file/optimistic-concurrency machinery
    find_and_replace_in_page and replace_last_block already have.

    Returns (ok, message, block_id, old_raw_content, new_raw_content).
    block_id, old_raw_content and new_raw_content are what undo/redo need:
    old_raw_content is the whole matched block's content exactly as it
    was before (undo restores this byte-for-byte, every paragraph and its
    formatting, not just the plain text of the one paragraph that
    changed); new_raw_content is the same block after splicing (what redo
    reapplies directly, without re-resolving anchor_text -- which might
    not even still be there, or might match differently, by the time
    someone redoes this).
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return False, "Could not read page content.", "", "", ""

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return False, f"Could not parse page XML: {e}", "", "", ""

    t_tag = f"{{{ONE_NS}}}T"
    candidates = []  # (block_id, raw_html) for every body OE whose text contains anchor_text
    for oe in _iter_body_oe(root):
        t_children = [c for c in list(oe) if c.tag == t_tag]
        if not t_children:
            continue
        raw_joined = "".join(t.text or "" for t in t_children)
        plain = html.unescape(re.sub(r"<[^>]+>", "", raw_joined))
        if anchor_text in plain:
            candidates.append((oe.get("objectID", ""), raw_joined))

    if not candidates:
        return False, (
            f"No block found containing {anchor_text!r}. Call read_live_page "
            f"first to find the exact wording."
        ), "", "", ""
    if len(candidates) > 1:
        return False, (
            f"Found {len(candidates)} different blocks containing {anchor_text!r} "
            f"-- ambiguous. Use more specific anchor_text that uniquely identifies "
            f"one paragraph."
        ), "", "", ""

    block_id, raw_html = candidates[0]
    if not block_id:
        return False, "The matching block has no stable ID (unexpected OneNote state) -- cannot safely target it.", "", "", ""

    new_paragraph_html = _sanitize_html_for_onenote(new_content)
    spliced, para_matches = _splice_paragraph_after(raw_html, anchor_text, new_paragraph_html)
    if spliced is None:
        return False, (
            f"Found {anchor_text!r} in {para_matches} different paragraphs within "
            f"the same block -- ambiguous. Use more specific anchor_text."
        ), "", "", ""

    ok, msg, _ = _com_replace_block_by_id(page_id, block_id, spliced, expected_raw=raw_html)
    return ok, msg, block_id, raw_html, spliced


def _com_list_pages(section_id: str) -> list[dict]:
    """List pages in a section via COM API. Returns list of {id, name}."""
    root = _com_get_hierarchy(4)
    if root is None:
        return []

    pages = []
    for sec in root.iter(f"{{{ONE_NS}}}Section"):
        if sec.get("ID") == section_id:
            for page in sec.iter(f"{{{ONE_NS}}}Page"):
                if page.get("isInRecycleBin") == "true":
                    continue
                pages.append({
                    "id": page.get("ID", ""),
                    "name": page.get("name", "(untitled)"),
                })
            break
    return pages


# ---------------------------------------------------------------------------
# MCP Write Tools
# ---------------------------------------------------------------------------


@mcp.tool()
async def list_live_notebooks() -> str:
    """List notebooks from the running OneNote app (live, not backup files).

    This uses the OneNote COM API and shows the notebooks currently open in
    the OneNote desktop app, including nested section groups (folders) and
    sections. Use this to find where to create, move, or delete things.
    """
    root = _com_get_hierarchy(3)
    if root is None:
        return "Could not connect to OneNote. Make sure the OneNote desktop app is installed."

    lines = []
    for nb in root.findall(f"{{{ONE_NS}}}Notebook"):
        nb_name = nb.get("name", "?")
        lines.append(f"\n## {nb_name}")
        lines.extend(_format_hierarchy_tree(nb, depth=1))

    if not lines:
        return "No notebooks found in OneNote."
    return "\n".join(lines)


@mcp.tool()
async def create_section(notebook_name: str, section_name: str) -> str:
    """Create a new section in a OneNote notebook (live, via COM API).

    Requires the OneNote desktop app to be installed.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name for the new section.
    """
    notebook_id = _com_find_notebook_id(notebook_name)
    if notebook_id is None:
        return (
            f"Could not find notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks."
        )

    ok, result = _com_create_section(notebook_id, section_name)
    if ok:
        _log_action(
            f"create_section | notebook={notebook_name} | created section '{section_name}'",
            {
                "type": "create_section", "object_id": result,
                "notebook_id": notebook_id, "section_name": section_name,
            },
        )
        return f"Section '{section_name}' created successfully (ID: {result})"
    return f"Failed to create section '{section_name}': {result}"


@mcp.tool()
async def create_section_group(notebook_name: str, group_name: str) -> str:
    """Create a new section group (folder) directly under a notebook.

    Requires the OneNote desktop app to be installed.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        group_name: Name for the new section group.
    """
    notebook_id = _com_find_notebook_id(notebook_name)
    if notebook_id is None:
        return (
            f"Could not find notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks."
        )

    ok, result = _com_create_section_group(notebook_id, group_name)
    if ok:
        _log_action(
            f"create_section_group | notebook={notebook_name} | created group '{group_name}'",
            {
                "type": "create_section_group", "object_id": result,
                "notebook_id": notebook_id, "group_name": group_name,
            },
        )
        return f"Section group '{group_name}' created successfully (ID: {result})"
    return f"Failed to create section group '{group_name}': {result}"


@mcp.tool()
async def move_section(notebook_name: str, section_name: str, destination_group_name: str = "") -> str:
    """Move a section to a different section group, or back to the notebook's top level.

    Also works to restore a deleted section out of the recycle bin -- pass
    the recycled section's name and the group (or "") you want it restored to.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name of the section to move (searched everywhere in the
            notebook, including the recycle bin).
        destination_group_name: Name of the destination section group. Leave
            empty ("") to move the section to the notebook's top level.
    """
    section_id = _com_find_section_id(notebook_name, section_name, include_recycled=True)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available sections."
        )

    if destination_group_name:
        destination_id = _com_find_section_group_id(notebook_name, destination_group_name)
        if destination_id is None:
            return (
                f"Could not find section group '{destination_group_name}' in notebook '{notebook_name}'. "
                f"Use list_live_notebooks to see available section groups."
            )
        destination_tag = "SectionGroup"
    else:
        destination_id = _com_find_notebook_id(notebook_name)
        if destination_id is None:
            return f"Could not find notebook '{notebook_name}'."
        destination_tag = "Notebook"

    old_parent = _com_find_section_parent(notebook_name, section_id)

    ok, result = _com_move_section(section_id, destination_id, destination_tag)
    if ok:
        dest_desc = destination_group_name or f"the top level of '{notebook_name}'"
        if old_parent:
            old_parent_id, old_parent_tag = old_parent
            _log_action(
                f"move_section | section={section_name} | moved to {dest_desc}",
                {
                    "type": "move_section",
                    "section_id": section_id,
                    "old_parent_id": old_parent_id,
                    "old_parent_tag": old_parent_tag,
                    "new_parent_id": destination_id,
                    "new_parent_tag": destination_tag,
                },
            )
        return f"Section '{section_name}' moved to {dest_desc}."
    return f"Failed to move section '{section_name}': {result}"


@mcp.tool()
async def rename_section(notebook_name: str, section_name: str, new_name: str) -> str:
    """Rename an existing section.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Current name of the section.
        new_name: New name for the section.
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available sections."
        )

    ok, result = _com_rename_section(section_id, new_name)
    if ok:
        _log_action(
            f"rename_section | section_id={section_id} | \"{section_name}\" -> \"{new_name}\"",
            {
                "type": "rename_section", "section_id": section_id,
                "old_name": section_name, "new_name": new_name,
            },
        )
        return f"Section '{section_name}' renamed to '{new_name}'."
    return f"Failed to rename section '{section_name}': {result}"


@mcp.tool()
async def rename_page(page_id: str, new_title: str) -> str:
    """Rename an existing page's title, leaving its content untouched.

    Args:
        page_id: The page ID (from list_live_pages).
        new_title: New title for the page.
    """
    old_title = _com_get_page_title(page_id)

    ok, result = _com_rename_page(page_id, new_title)
    if ok:
        if old_title is not None:
            _log_action(
                f"rename_page | page_id={page_id} | \"{old_title}\" -> \"{new_title}\"",
                {
                    "type": "rename_page", "page_id": page_id,
                    "old_title": old_title, "new_title": new_title,
                },
            )
        return f"Page renamed to '{new_title}'."
    return f"Failed to rename page: {result}"


@mcp.tool()
async def delete_section(notebook_name: str, section_name: str, confirm: bool = False) -> str:
    """Delete a section (moves it to the OneNote recycle bin, not permanent).

    This is a two-step, irreversible-feeling operation: call it once with
    confirm left as false to get a preview of exactly what would be deleted,
    then call it again with confirm=true only after the user has explicitly
    agreed to delete that specific section.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name of the section to delete.
        confirm: Must be explicitly set to true to actually perform the
            deletion. Defaults to false, which only returns a preview.
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available sections."
        )

    if not confirm:
        pages = _com_list_pages(section_id)
        return (
            f"PREVIEW (nothing deleted yet): this would move section '{section_name}' "
            f"in notebook '{notebook_name}' ({len(pages)} page(s)) to the OneNote recycle "
            f"bin. Ask the user to explicitly confirm this exact section before calling "
            f"again with confirm=true."
        )

    old_parent = _com_find_section_parent(notebook_name, section_id)

    ok, result = _com_delete_hierarchy(section_id)
    if ok:
        if old_parent:
            old_parent_id, old_parent_tag = old_parent
            _log_action(
                f"delete_section | notebook={notebook_name} | deleted section '{section_name}'",
                {
                    "type": "delete_section",
                    "section_id": section_id,
                    "old_parent_id": old_parent_id,
                    "old_parent_tag": old_parent_tag,
                },
            )
        return f"Section '{section_name}' deleted (moved to recycle bin)."
    return f"Failed to delete section '{section_name}': {result}"


@mcp.tool()
async def delete_page(page_id: str, confirm: bool = False) -> str:
    """Delete a page (moves it to the OneNote recycle bin, not permanent).

    This is a two-step, irreversible-feeling operation: call it once with
    confirm left as false to get a preview, then call it again with
    confirm=true only after the user has explicitly agreed to delete that
    specific page.

    Args:
        page_id: The page ID (from list_live_pages).
        confirm: Must be explicitly set to true to actually perform the
            deletion. Defaults to false, which only returns a preview.
    """
    if not confirm:
        return (
            f"PREVIEW (nothing deleted yet): this would move page (ID: {page_id}) to "
            f"the OneNote recycle bin. Ask the user to explicitly confirm this exact "
            f"page before calling again with confirm=true."
        )

    ok, result = _com_delete_hierarchy(page_id)
    if ok:
        _log_action(
            f"delete_page | page_id={page_id} | deleted (not automatically undoable)",
            {"type": "delete_page", "page_id": page_id},
        )
        return "Page deleted (moved to recycle bin)."
    return f"Failed to delete page: {result}"


@mcp.tool()
async def list_recycle_bin(notebook_name: str) -> str:
    """List sections and pages currently in a notebook's recycle bin.

    Use this to find the exact name of something to restore with
    move_section, or to double-check what a delete_section call actually
    removed.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
    """
    root = _com_get_hierarchy(4)
    if root is None:
        return "Could not connect to OneNote. Make sure the OneNote desktop app is installed."

    lines = []
    for nb in root.iter(f"{{{ONE_NS}}}Notebook"):
        if nb.get("name", "").lower() != notebook_name.lower():
            continue
        for sec in nb.iter(f"{{{ONE_NS}}}Section"):
            if sec.get("isInRecycleBin") == "true":
                lines.append(f"- [section] {sec.get('name', '?')}")
                continue
            for page in sec.iter(f"{{{ONE_NS}}}Page"):
                if page.get("isInRecycleBin") == "true":
                    lines.append(f"- [page] {page.get('name', '?')}  (in section '{sec.get('name', '?')}', id: {page.get('ID', '')})")

    if not lines:
        return f"Recycle bin for '{notebook_name}' is empty."
    return "\n".join(lines)


@mcp.tool()
async def create_page(notebook_name: str, section_name: str, title: str, content: str) -> str:
    """Create a new page in a OneNote notebook section.

    The content is written as HTML. You can use basic HTML tags like
    <b>, <i>, <br>, <ul>, <li>, <h1>-<h6>, etc.

    Requires the OneNote desktop app to be installed.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name of the section within the notebook.
        title: Title for the new page.
        content: The page content (plain text or HTML).
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks and sections."
        )

    ok, msg = _com_create_page(section_id, title, content)
    if ok:
        id_match = re.search(r"\(ID: (.+)\)$", msg)
        if id_match:
            _log_action(
                f"create_page | notebook={notebook_name} | section={section_name} | created page '{title}'",
                {
                    "type": "create_page", "object_id": id_match.group(1),
                    "section_id": section_id, "title": title, "content": content,
                },
            )
    return msg


@mcp.tool()
async def list_live_pages(notebook_name: str, section_name: str) -> str:
    """List pages in a section from the running OneNote app.

    Use this to find page IDs for appending content to existing pages.

    Args:
        notebook_name: Name of the notebook.
        section_name: Name of the section.
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks and sections."
        )

    pages = _com_list_pages(section_id)
    if not pages:
        return "No pages found in this section."

    lines = []
    for p in pages:
        lines.append(f"- {p['name']}  (id: {p['id']})")
    return "\n".join(lines)


@mcp.tool()
async def read_live_page(page_id: str) -> str:
    """Read the full text content of a single page from the running OneNote app (live).

    Unlike read_section (which reads from a local backup snapshot and can be
    outdated), this fetches the current content directly from the OneNote
    desktop app via its COM API.

    Args:
        page_id: The page ID (from list_live_pages).
    """
    xml_content = _com_get_page_content(page_id)
    if xml_content is None:
        return (
            "Could not read page. Make sure the OneNote desktop app is "
            "installed and the page ID is correct (use list_live_pages)."
        )
    text = _extract_text_from_page_xml(xml_content)
    return text or "(page has no text content)"


@mcp.tool()
async def read_live_section(notebook_name: str, section_name: str) -> str:
    """Read the full text content of every page in a section from the running OneNote app (live).

    Unlike read_section (which reads from a local backup snapshot and can be
    outdated or missing recently added pages/sections entirely), this reads
    current content directly from the OneNote desktop app via its COM API.

    Args:
        notebook_name: Name of the notebook (from list_live_notebooks).
        section_name: Name of the section within the notebook.
    """
    section_id = _com_find_section_id(notebook_name, section_name)
    if section_id is None:
        return (
            f"Could not find section '{section_name}' in notebook '{notebook_name}'. "
            f"Use list_live_notebooks to see available notebooks and sections."
        )

    pages = _com_list_pages(section_id)
    if not pages:
        return "No pages found in this section."

    parts = []
    for p in pages:
        xml_content = _com_get_page_content(p["id"])
        text = _extract_text_from_page_xml(xml_content) if xml_content else ""
        parts.append(f"## {p['name']}\n{text or '(no text content)'}")

    return "\n\n".join(parts)


@mcp.tool()
async def insert_image_from_file(page_id: str, file_path: str) -> str:
    """Insert an image onto an existing page by reading it from a local file.

    Only a path needs to be passed here, never the image data itself, which
    keeps large images out of the conversation entirely.

    file_path can be a full absolute path, or just a bare filename (e.g.
    "Background.png") -- a bare filename is looked up in the user's default
    image drop folder (Downloads, unless overridden), so a user can just
    save a file there and refer to it by name.

    Requires the OneNote desktop app to be installed.

    Args:
        page_id: The page ID (from list_live_pages).
        file_path: Absolute path to the image file, or just its filename if
            it's in the default drop folder (png, jpeg, gif, etc.).
    """
    ok, msg = _com_insert_image_from_file(page_id, file_path)
    if ok:
        _log_action(
            f"insert_image_from_file | page_id={page_id} | inserted image from {file_path} (not automatically undoable)",
            {"type": "insert_image_from_file", "page_id": page_id, "file_path": file_path},
        )
    return msg


@mcp.tool()
async def append_to_page(page_id: str, content: str) -> str:
    """Append content to an existing OneNote page.

    The content is added as a new outline block at the bottom of the page.
    Supports HTML formatting (<b>, <i>, <br>, <ul>, <li>, etc.).

    Args:
        page_id: The page ID (from list_live_pages).
        content: The content to append (plain text or HTML).
    """
    ok, msg, new_outline_id = _com_append_to_page(page_id, content)
    if ok:
        _log_action(
            f"append_to_page | page_id={page_id} | appended a new block",
            {
                "type": "append_to_page", "page_id": page_id,
                "outline_id": new_outline_id, "content": content,
            },
        )
    return msg


@mcp.tool()
async def insert_block_after(page_id: str, anchor_text: str, new_content: str) -> str:
    """Insert new_content as a new paragraph immediately after the
    existing text containing anchor_text -- lands right there in the
    page's reading flow, whether that text is its own block or part of a
    bigger one (e.g. a page written by a single create_page call, whose
    whole body is technically one block with several paragraphs inside
    it -- this still lands between the right two paragraphs, not at the
    end of the page).

    Use this for putting something in a specific place partway through a
    page -- e.g. "add this paragraph right after the one about X". For
    adding to the very end of the page, use append_to_page instead; for
    changing something that's already there, use replace_last_block or
    find_and_replace_in_page.

    anchor_text must match exactly (case-sensitive, no fuzzy/approximate
    matching), same contract as find_and_replace_in_page -- if in doubt,
    call read_live_page first to see the exact current wording, then pass
    back the exact text you found there. If anchor_text matches more than
    one paragraph on the page, or none, this is reported honestly rather
    than guessed at -- narrow it down to something that uniquely
    identifies one paragraph first.

    Args:
        page_id: The page ID (from list_live_pages).
        anchor_text: Exact text identifying which existing paragraph the
            new content should be inserted after.
        new_content: The content for the new paragraph (plain text or HTML).
    """
    ok, msg, block_id, old_raw, new_raw = _com_insert_block_after(page_id, anchor_text, new_content)
    if ok:
        _log_action(
            f"insert_block_after | page_id={page_id} | inserted a new paragraph after \"{anchor_text}\"",
            {
                "type": "insert_block_after", "page_id": page_id, "block_id": block_id,
                "old_text": old_raw, "new_text": new_raw,
                "anchor_text": anchor_text, "inserted_content": new_content,
            },
        )
    return msg


@mcp.tool()
async def replace_last_block(page_id: str, new_content: str, confirm: bool = False) -> str:
    """Replace the last content block on a page with new_content, without
    needing to know or repeat what it currently says.

    Use this for "no, change what you just wrote" -- it always targets the
    most recently written block on the page (typically the one just added
    by append_to_page), regardless of what it contained. For editing
    something older or deeper in a page, use find_and_replace_in_page
    instead.

    Pages built from a single create_page call put the entire body into
    one single block, so there "the last block" IS the whole page body --
    this would replace everything, not just a trailing sentence. That
    case is treated as an irreversible-feeling operation the same way
    delete_section/delete_page are: call once with confirm left as false
    to get a preview, then again with confirm=true only after the user
    has explicitly agreed to replace this page's entire content. On a
    page with more than one block, no confirmation is needed -- only the
    last, narrower block is ever affected there.

    Args:
        page_id: The page ID (from list_live_pages).
        new_content: The new content for that block (plain text or HTML).
        confirm: Only checked when the page turns out to have just one
            content block. Defaults to false, which returns a preview
            (and changes nothing) in that case.
    """
    if not confirm:
        block_count = _com_count_body_text_blocks(page_id)
        if block_count == 1:
            return (
                "PREVIEW (nothing replaced yet): this page has only one content "
                "block, so replace_last_block would overwrite its ENTIRE body, "
                "not just a trailing part of it. Ask the user to explicitly "
                "confirm that before calling again with confirm=true."
            )

    ok, msg, old_text, block_id = _com_replace_last_block(page_id, new_content)
    if ok:
        _log_action(
            f"replace_last_block | page_id={page_id} | replaced the last block",
            {
                "type": "replace_last_block", "page_id": page_id,
                "old_text": old_text, "new_text": new_content, "block_id": block_id,
            },
        )
    return msg


@mcp.tool()
async def find_and_replace_in_page(
    page_id: str, find_text: str, replace_text: str, replace_all: bool = False
) -> str:
    """Find and replace a specific piece of text within an existing page.

    find_text must match exactly (case-sensitive, no fuzzy/approximate
    matching) -- if in doubt, call read_live_page first to see the exact
    current wording, then pass back the exact text you found there. If the
    page's content came from a loosely-worded request (e.g. "change the
    part about X"), identify the precise passage yourself from what
    read_live_page returns, and ask the user to clarify if it's ambiguous
    (e.g. it appears more than once) -- don't guess.

    Args:
        page_id: The page ID (from list_live_pages).
        find_text: The exact text to search for.
        replace_text: The text to replace it with.
        replace_all: If true, replace every occurrence found; if false
            (default), replace only the first one found.
    """
    ok, msg, count = _com_find_and_replace_in_page(page_id, find_text, replace_text, replace_all)
    if ok and count > 0:
        _log_action(
            f"find_and_replace_in_page | page_id={page_id} | replaced \"{find_text}\" with \"{replace_text}\" ({count}x)",
            {
                "type": "find_and_replace_in_page",
                "page_id": page_id,
                "old_text": find_text,
                "new_text": replace_text,
                "replace_all": replace_all,
            },
        )
    return msg


def _reverse_logged_action(undo_data: dict) -> tuple[bool, str]:
    """Reverse one logged action, based on its recorded undo_data. Returns
    (ok, message). Some action types (delete_page, insert_image_from_file)
    are intentionally not automatically reversible -- see history notes."""
    action_type = undo_data.get("type")

    if action_type == "rename_section":
        return _com_rename_section(undo_data["section_id"], undo_data["old_name"])

    if action_type == "rename_page":
        return _com_rename_page(undo_data["page_id"], undo_data["old_title"])

    if action_type in ("move_section", "delete_section"):
        # Undoing a delete means moving the section back out of the recycle
        # bin to where it was -- same underlying operation as undoing a move.
        return _com_move_section(
            undo_data["section_id"], undo_data["old_parent_id"], undo_data["old_parent_tag"]
        )

    if action_type in ("create_section", "create_section_group", "create_page"):
        return _com_delete_hierarchy(undo_data["object_id"])

    if action_type == "find_and_replace_in_page":
        ok, msg, _ = _com_find_and_replace_in_page(
            undo_data["page_id"], undo_data["new_text"], undo_data["old_text"],
            replace_all=undo_data.get("replace_all", False),
        )
        return ok, msg

    if action_type == "replace_last_block":
        block_id = undo_data.get("block_id")
        if block_id:
            # Targets the exact block this action changed, immune to the
            # page having grown new blocks since (see _com_replace_block_by_id).
            ok, msg, _ = _com_replace_block_by_id(undo_data["page_id"], block_id, undo_data["old_text"])
            return ok, msg
        # Fallback for entries logged before block-ID tracking existed --
        # re-resolves "the last block" fresh, same imprecision this
        # always had before today's fix, not made any worse.
        ok, msg, _, _ = _com_replace_last_block(undo_data["page_id"], undo_data["old_text"])
        return ok, msg

    if action_type == "append_to_page":
        outline_id = undo_data.get("outline_id")
        if outline_id:
            return _com_delete_page_content_object(undo_data["page_id"], outline_id)
        # Fallback for older log entries logged before outline_id was captured.
        return _com_remove_last_block(undo_data["page_id"])

    if action_type == "insert_block_after":
        block_id = undo_data.get("block_id")
        if block_id and "old_text" in undo_data:
            # Restores the whole block's content byte-for-byte (every
            # paragraph, not just the one that changed) -- this edits the
            # existing block in place rather than deleting a separately
            # created object, since insert_block_after never creates a
            # new block/outline, it splices into an existing one.
            ok, msg, _ = _com_replace_block_by_id(undo_data["page_id"], block_id, undo_data["old_text"])
            return ok, msg
        return False, "This entry has no stable block ID recorded -- cannot safely undo it automatically."

    if action_type in ("delete_page", "insert_image_from_file"):
        return False, (
            f"'{action_type}' cannot be automatically undone -- restore it manually "
            f"in OneNote if needed."
        )

    return False, f"Unknown action type in log: {action_type!r}"


@mcp.tool()
async def get_history_file_path() -> str:
    """Return where the action history log actually lives on disk.

    list_recent_actions only shows a limited, reformatted view -- use this
    when you (or the user) want to open history.md directly, e.g. to read
    the full log in a text editor, or to check history.archive.md for
    older entries that have rotated out of the active log.
    """
    lines = [f"Active log: {HISTORY_FILE}"]
    if HISTORY_FILE.exists():
        lines.append(f"  ({len(_read_history_entries())} entries)")
    else:
        lines.append("  (doesn't exist yet -- no actions logged so far)")
    lines.append(f"Archive (older, rotated-out entries): {HISTORY_ARCHIVE_FILE}")
    lines.append(
        "  (exists)" if HISTORY_ARCHIVE_FILE.exists() else "  (doesn't exist yet)"
    )
    return "\n".join(lines)


@mcp.tool()
async def list_recent_actions(limit: int = 10) -> str:
    """List the most recently logged actions (creates, renames, moves,
    replacements, deletes, etc.), newest first.

    Use this to see what actually happened before deciding whether to call
    undo_last_action, or to figure out exactly what to fix yourself if the
    wrong thing was changed.

    Args:
        limit: Maximum number of entries to show (default 10).
    """
    entries = _read_history_entries()
    if not entries:
        return "No actions logged yet."

    recent = list(reversed(entries[-limit:]))
    lines = []
    for entry in recent:
        parsed = _parse_history_line(entry)
        marker = " [undone]" if parsed["undone"] else ""
        lines.append(f"- {parsed['timestamp']} | {parsed['summary']}{marker}")
    return "\n".join(lines)


# English labels for the raw undo_data keys get_action_detail dumps --
# the keys themselves are internal (page_id, old_text, ...), these are
# what a human should actually see.
_ACTION_FIELD_LABELS = {
    "page_id": "Page ID",
    "section_id": "Section ID",
    "notebook_id": "Notebook ID",
    "object_id": "Object ID",
    "outline_id": "Outline ID",
    "old_name": "Old name",
    "new_name": "New name",
    "old_title": "Old title",
    "new_title": "New title",
    "old_text": "Old",
    "new_text": "New",
    "title": "Title",
    "content": "Content",
    "section_name": "Section name",
    "group_name": "Group name",
    "old_parent_id": "Old parent ID",
    "old_parent_tag": "Old parent type",
    "new_parent_id": "New parent ID",
    "new_parent_tag": "New parent type",
    "replace_all": "Replaced all occurrences",
    "anchor_text": "Inserted after (text)",
    "block_id": "Block ID",
    "inserted_content": "Inserted paragraph",
    "file_path": "File path",
}


@mcp.tool()
async def get_action_detail(n: int = 1) -> str:
    """Return the complete, unabridged data for one logged action -- not
    the one-line summary list_recent_actions shows.

    Args:
        n: Which entry to fetch, counting back from the most recent
            (1 = the last action, 2 = the one before that, ...). Matches
            the order list_recent_actions lists entries in.

    How to present this to the user: as a small table/detail card, one
    row per field, English labels throughout (this project is for
    general use, not tied to any one person's language) -- roughly
    Timestamp, Action, Page, then whichever old/new or other
    type-specific fields are present (already labeled below), in that
    order. Leave out "Replaced all occurrences" -- how many matches got
    replaced is rarely useful in this view, even though it's in the raw
    data.

    Add two rows yourself, written from the surrounding conversation, not
    from this tool's output -- it has no way to know either:
    - Reason: a short sentence on why this change was likely made.
    - Summary: one sentence describing what the new value actually says.

    If a page_id or section_id shows up, resolving it to a human-readable
    name/path first (e.g. via list_live_pages or list_live_notebooks)
    makes the table much more readable -- worth the extra call when
    practical, but keep the raw ID visible too, since that's what other
    tools here actually need.
    """
    entries = _read_history_entries()
    if not entries:
        return "No actions logged yet."
    if n < 1 or n > len(entries):
        return f"No entry at position {n} -- there are {len(entries)} logged action(s) (use list_recent_actions to see them)."

    parsed = _parse_history_line(list(reversed(entries))[n - 1])
    undo_data = dict(parsed["undo_data"])
    action_type = undo_data.pop("type", "unknown")

    lines = [
        f"Timestamp: {parsed['timestamp']}",
        f"Action: {action_type}",
        f"Log summary: {parsed['summary']}",
        f"Currently undone: {parsed['undone']}",
    ]
    for key, value in undo_data.items():
        label = _ACTION_FIELD_LABELS.get(key, key)
        lines.append(f"{label}: {value}")
    return "\n".join(lines)


@mcp.tool()
async def undo_last_action() -> str:
    """Undo the most recent logged action.

    Reverses whichever action is most recent in the history log (rename,
    move, delete, find/replace, replace_last_block, append, or create) and
    marks it as undone in the log. Deleting a page and inserting an image
    cannot be automatically undone -- if the last action was one of those,
    this reports that plainly instead of silently doing nothing or undoing
    something older instead. Calling this repeatedly walks further back
    through the log, one action at a time; redo_last_action reverses that.
    """
    last = _find_last_undoable_action()
    if last is None:
        return "Nothing to undo (no actions logged yet, or everything already undone)."

    ok, msg = _reverse_logged_action(last["undo_data"])
    if ok:
        _set_action_undone_flag(last["id"], last["raw"], True)
        return f"Undid: {last['timestamp']} | {last['summary']}. ({msg})"
    return f"Could not undo '{last['timestamp']} | {last['summary']}': {msg}"


def _apply_logged_action(data: dict) -> tuple[bool, str, dict]:
    """Reapply one logged action forward -- the counterpart to
    _reverse_logged_action, used by redo_last_action.

    Some fields needed to redo an action (e.g. the name it was renamed TO,
    or the content a block was replaced WITH) aren't needed for undo, so
    they're only present on entries logged after redo support was added.
    An older entry missing them is reported honestly as not redoable,
    rather than guessed at or silently skipped.

    Returns (ok, message, updates). updates is a dict of fields the caller
    should merge into the history entry afterwards -- needed specifically
    for the create_* and append_to_page types, whose undo (deleting the
    object) means redo has to *recreate* it, which gets a brand new
    object/outline ID from OneNote, never the original one. Without
    updating the entry's stored ID to that new one, a second undo of the
    same entry would try to delete the long-gone original ID instead of
    the thing that's actually on the page now.
    """
    action_type = data.get("type")

    if action_type == "rename_section":
        if "new_name" not in data:
            return False, "This entry predates redo support (no new name recorded) -- redo it by hand.", {}
        ok, msg = _com_rename_section(data["section_id"], data["new_name"])
        return ok, msg, {}

    if action_type == "rename_page":
        if "new_title" not in data:
            return False, "This entry predates redo support (no new title recorded) -- redo it by hand.", {}
        ok, msg = _com_rename_page(data["page_id"], data["new_title"])
        return ok, msg, {}

    if action_type == "move_section":
        if "new_parent_id" not in data:
            return False, "This entry predates redo support (no destination recorded) -- redo it by hand.", {}
        ok, msg = _com_move_section(data["section_id"], data["new_parent_id"], data["new_parent_tag"])
        return ok, msg, {}

    if action_type == "delete_section":
        # Redoing a delete just means deleting it again -- section_id alone
        # is enough, nothing extra to have captured at log time, and
        # deleting doesn't mint a new ID the way recreating does.
        ok, msg = _com_delete_hierarchy(data["section_id"])
        return ok, msg, {}

    if action_type == "create_section":
        if "notebook_id" not in data or "section_name" not in data:
            return False, "This entry predates redo support (no recreation data recorded) -- redo it by hand.", {}
        ok, result = _com_create_section(data["notebook_id"], data["section_name"])
        if ok:
            return True, f"Section recreated (new ID: {result})", {"object_id": result}
        return False, result, {}

    if action_type == "create_section_group":
        if "notebook_id" not in data or "group_name" not in data:
            return False, "This entry predates redo support (no recreation data recorded) -- redo it by hand.", {}
        ok, result = _com_create_section_group(data["notebook_id"], data["group_name"])
        if ok:
            return True, f"Section group recreated (new ID: {result})", {"object_id": result}
        return False, result, {}

    if action_type == "create_page":
        if "section_id" not in data or "title" not in data or "content" not in data:
            return False, "This entry predates redo support (no recreation data recorded) -- redo it by hand.", {}
        ok, msg = _com_create_page(data["section_id"], data["title"], data["content"])
        if not ok:
            return False, msg, {}
        id_match = re.search(r"\(ID: (.+)\)$", msg)
        if not id_match:
            # The page WAS created -- ok is True -- but its new ID
            # couldn't be scraped back out of the success message, so
            # don't claim {"object_id": ...} in updates: that would leave
            # the entry's stale, already-deleted original ID in place
            # silently, same failure mode as append_to_page's outline_id
            # capture above. Surface it instead of hiding it.
            return True, f"{msg} (could not confirm the new page's exact ID -- a later undo of this entry may not target it)", {}
        return True, msg, {"object_id": id_match.group(1)}

    if action_type == "find_and_replace_in_page":
        # Already fully bidirectional from the start -- old_text/new_text
        # are both stored for undo's sake, so redo just reapplies them as-is.
        # Edits an existing block in place, so no new ID to track.
        ok, msg, _ = _com_find_and_replace_in_page(
            data["page_id"], data["old_text"], data["new_text"],
            replace_all=data.get("replace_all", False),
        )
        return ok, msg, {}

    if action_type == "replace_last_block":
        if "new_text" not in data:
            return False, "This entry predates redo support (no new content recorded) -- redo it by hand.", {}
        block_id = data.get("block_id")
        if block_id:
            ok, msg, _ = _com_replace_block_by_id(data["page_id"], block_id, data["new_text"])
            return ok, msg, {}
        ok, msg, _, _ = _com_replace_last_block(data["page_id"], data["new_text"])
        return ok, msg, {}

    if action_type == "append_to_page":
        if "content" not in data:
            return False, "This entry predates redo support (no content recorded) -- redo it by hand.", {}
        ok, msg, new_outline_id = _com_append_to_page(data["page_id"], data["content"])
        if not ok:
            return False, msg, {}
        if not new_outline_id:
            # The append itself succeeded, but _com_append_to_page's own
            # follow-up fetch to read back the new block's ID came up
            # empty (a transient COM/fetch hiccup, not a failed append).
            # Don't let that empty value overwrite whatever ID was
            # already on this entry -- entry.update(extra) would clobber
            # a perfectly good, specific outline_id with "", and a later
            # undo of THIS entry would then fall back to "delete whatever
            # is currently last on the page" instead of a known ID --
            # unsafe if anything else got appended in between. Keeping
            # the stale ID means that fallback undo instead fails loudly
            # against a nonexistent ID, which is the safe failure mode.
            return True, f"{msg} (could not confirm the new block's exact ID -- a later undo of this entry may not target it precisely)", {}
        return True, msg, {"outline_id": new_outline_id}

    if action_type == "insert_block_after":
        block_id = data.get("block_id")
        if block_id and "new_text" in data:
            # Reapplies the exact same already-spliced content directly --
            # no need to re-resolve anchor_text (which might not even
            # match anything, or match differently, by redo time).
            ok, msg, _ = _com_replace_block_by_id(data["page_id"], block_id, data["new_text"])
            return ok, msg, {}
        if "anchor_text" in data and "inserted_content" in data:
            # Fallback for entries logged before block-ID tracking
            # existed -- re-resolves anchor_text fresh.
            ok, msg, new_block_id, _, new_raw = _com_insert_block_after(
                data["page_id"], data["anchor_text"], data["inserted_content"]
            )
            if not ok:
                return False, msg, {}
            return True, msg, {"block_id": new_block_id, "new_text": new_raw}
        return False, "This entry predates redo support (no recreation data recorded) -- redo it by hand.", {}

    if action_type in ("delete_page", "insert_image_from_file"):
        return False, f"'{action_type}' was never undone in the first place, so there's nothing to redo.", {}

    return False, f"Unknown action type in log: {action_type!r}", {}


@mcp.tool()
async def redo_last_action() -> str:
    """Redo the most recently undone action (the counterpart to
    undo_last_action).

    Finds whichever undone action was undone most recently and reapplies
    it, regardless of how far back in the log it sits -- so undoing three
    actions in a row and then calling this three times puts all three back,
    in the right order. Entries logged before redo support existed don't
    carry the data needed to redo them and are reported honestly rather
    than guessed at. Deleting a page or inserting an image was never
    undoable to begin with, so there's nothing for this to redo there.
    """
    last = _find_last_redoable_action()
    if last is None:
        return "Nothing to redo (no actions have been undone, or everything undone has already been redone)."

    ok, msg, updates = _apply_logged_action(last["undo_data"])
    if ok:
        _set_action_undone_flag(last["id"], last["raw"], False, extra=updates)
        return f"Redid: {last['timestamp']} | {last['summary']}. ({msg})"
    return f"Could not redo '{last['timestamp']} | {last['summary']}': {msg}"


@mcp.tool()
async def list_drop_folder_images(limit: int = 20) -> str:
    """List image files in the default image drop folder (see
    insert_image_from_file), sorted newest first.

    Use this to see what's actually available before calling
    insert_image_from_file with a bare filename -- e.g. to find "the last
    few files downloaded" instead of needing an already-known exact name.

    Args:
        limit: Maximum number of files to list (default 20).
    """
    if not IMAGE_DROP_DIR.is_dir():
        return f"Image drop folder not found: {IMAGE_DROP_DIR}"

    image_extensions = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
    files = [
        f for f in IMAGE_DROP_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in image_extensions
    ]
    if not files:
        return f"No image files found in {IMAGE_DROP_DIR}"

    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    lines = []
    for f in files[:limit]:
        stat = f.stat()
        size_kb = stat.st_size / 1024
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
        lines.append(f"- {f.name}  ({size_kb:.0f} KB, modified {modified})")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if not ONENOTE_DIR.exists():
        log.error(
            "OneNote backup directory not found: %s\n"
            "Set the ONENOTE_BACKUP_DIR environment variable to the correct path.",
            ONENOTE_DIR,
        )
        sys.exit(1)

    log.info("Starting OneNote MCP server (local files)...")
    log.info("Reading from: %s", ONENOTE_DIR)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
