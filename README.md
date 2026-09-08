# OneNote (Local) — MCP Server

An [MCP (Model Context Protocol)](https://modelcontextprotocol.io/) server that gives Claude direct access to your local Microsoft OneNote notebooks — no Microsoft account, API key, or Azure app registration needed. Everything runs locally on your own machine.

It can read from OneNote's local backup snapshots (fast) **and** talk directly to the running OneNote desktop app via its COM automation interface for always-current reads and for writing (creating sections, pages, and appending content).

## Background

This started as [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp), a small MCP server for reading local OneNote backup files and writing to OneNote via COM. Rail1337 runs a D&D campaign and keeps all the session notes, quizzes, and the DM diary in OneNote — and got tired of copy-pasting pages into Claude by hand just to ask "what happened in session 4 again?". The original project covered basic reading and page creation, but couldn't read a single page or a whole section straight from the live app, and couldn't create new sections at all. So it got picked up and substantially extended (see below) to close those gaps. This stays open source.

## What this version adds on top of the original

- **`read_live_page` and `read_live_section`** — entirely new. The original could only create/append pages live; there was no way to read a single page or a whole section directly from the running app, only from the (potentially outdated) local backup. This closes that gap.
- **`create_section`** — entirely new. The original could only create pages inside sections that already existed.
- **Fixed a page-title bug**: the original escaped `]]>` in page *bodies* but not in page *titles*, so a title containing that sequence could break the CDATA wrapper and corrupt the page's XML. Titles are now escaped the same way.
- **Language-independent backup detection**: the original hardcoded the English folder name `Backup`. It's now auto-detected regardless of Windows/Office display language (e.g. German "Sicherung"), by scanning for whichever folder actually contains `.one` files.
- Packaged as a self-contained `.mcpb` bundle for one-click installation in Claude Desktop (no manual Python/`uv` setup).

## Tools

### Reading from local backup snapshots
These are fast, but only as fresh as OneNote's last automatic backup — recently added sections or pages may not show up yet.

| Tool | Description |
|------|-------------|
| `list_notebooks` | List all notebooks found in the local backup |
| `list_sections` | List sections in a notebook |
| `read_section` | Read all text content of a section |
| `search_notes` | Search text across every notebook and section |
| `list_all_sections` | Full overview of every section in every notebook |
| `get_notebook_summary` | Notebook overview with content previews |

### Live, via the running OneNote app
These talk to the OneNote desktop app directly, so they're always current — but each call is a bit slower since it drives OneNote via COM automation.

| Tool | Description |
|------|-------------|
| `list_live_notebooks` | List notebooks/sections currently open in OneNote |
| `create_section` | Create a brand-new section in a notebook |
| `create_page` | Create a new page in a section |
| `list_live_pages` | List pages in a section, with IDs for reading/appending |
| `read_live_page` | Read the full text of a single page |
| `read_live_section` | Read every page in a section |
| `append_to_page` | Append content to an existing page |

## Prerequisites

- Windows, with the Microsoft OneNote desktop app installed
- For the backup-based reading tools: OneNote must have run at least once so a local backup exists
- For the live tools: OneNote just needs to be installed — its COM API starts it automatically if it isn't already running

## Installation

Drag this `.mcpb` file onto Claude Desktop (or double-click it) and confirm the install prompt. Claude Desktop manages the Python/`uv` runtime itself, so nothing else needs to be installed. Restart Claude Desktop afterwards.

## How it works

- The backup-reading tools parse the `.one` files OneNote automatically keeps under `%LOCALAPPDATA%\Microsoft\OneNote\16.0\<Backup>\`. The folder name is auto-detected regardless of Windows language (English "Backup", German "Sicherung", etc.) — it scans for whichever subfolder actually contains `.one` files.
- The live tools drive the OneNote desktop app through its COM automation interface (`OneNote.Application`), invoked via PowerShell under the hood.

## Credits & License

Originally created by [mhzarem](https://github.com/mhzarem/onenote-mcp) and released under the MIT license. Rail1337 is **not** the original author — this is a modified, substantially extended derivative of that project. Stays open source under the same MIT license.

MIT — see the original project's license terms.
