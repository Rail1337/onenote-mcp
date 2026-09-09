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
import logging
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
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

    if not image_format:
        return False, f"Could not determine image format for: {file_path}"

    return _com_insert_image_bytes(page_id, raw_bytes, image_format)


def _com_append_to_page(page_id: str, body_html: str) -> tuple[bool, str]:
    """Append content to an existing page using the OneNote COM API."""
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
        if ok:
            return True, "Content appended successfully."
        return False, f"Failed to append content: {output}"
    finally:
        try:
            os.remove(body_file)
        except OSError:
            pass


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

    ok, result = _com_move_section(section_id, destination_id, destination_tag)
    if ok:
        dest_desc = destination_group_name or f"the top level of '{notebook_name}'"
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
        return f"Section '{section_name}' renamed to '{new_name}'."
    return f"Failed to rename section '{section_name}': {result}"


@mcp.tool()
async def rename_page(page_id: str, new_title: str) -> str:
    """Rename an existing page's title, leaving its content untouched.

    Args:
        page_id: The page ID (from list_live_pages).
        new_title: New title for the page.
    """
    ok, result = _com_rename_page(page_id, new_title)
    if ok:
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

    ok, result = _com_delete_hierarchy(section_id)
    if ok:
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
    ok, msg = _com_append_to_page(page_id, content)
    return msg


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
