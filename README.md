I rebuilt this local MCP server so Claude can read and write my local Microsoft OneNote notebooks directly — no Microsoft account, API key, or Azure app registration needed, everything runs on my own machine. It reads from OneNote's local backup snapshots for quick browsing and search, and talks directly to the running OneNote desktop app via its COM automation interface for always-current reads and for writing and organizing content.

## Story

I run a D&D campaign, and years of session notes, NPC write-ups, quizzes, and my DM diary all live in OneNote. I didn't want to migrate all of that to Notion or one of the other note apps that already has an official Claude connector — OneNote is simply where the notes already were, and restructuring years of notebooks just to get a connector felt like the wrong trade-off. But I also didn't want to keep manually copy-pasting pages into Claude every time I needed context — "wait, what actually happened in session 4?" shouldn't mean digging through OneNote by hand every time. So instead of switching tools, I built the bridge myself.

This started as [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp), a small MCP server for reading local OneNote backup files and writing to OneNote via COM. It covered the basics, but couldn't read a single page or a whole section straight from the live app, couldn't organize sections at all, and had a small escaping bug in page titles. I picked it up and substantially extended it to close those gaps. It stays open source.

## Added Features

| Type | Feature | Description |
|---|---|---|
| Reading | `read_live_page` | Read the full text of a single page directly from the running OneNote app |
| Reading | `read_live_section` | Read every page in a section directly from the running app |
| Reading | `list_recycle_bin` | List sections and pages currently in a notebook's recycle bin |
| Organizing | `create_section` | Create a new section in a notebook |
| Organizing | `create_section_group` | Create a new section group (folder) in a notebook |
| Organizing | `move_section` | Move a section to a different section group, or back to the notebook's top level — also restores a section out of the recycle bin |
| Organizing | `rename_section` | Rename an existing section |
| Organizing | `rename_page` | Rename an existing page's title, without touching its content |
| Deleting | `delete_section` | Delete a section (moved to the OneNote recycle bin, not permanent) — requires an explicit confirmation before it executes |
| Deleting | `delete_page` | Delete a page — same recycle-bin safety net and confirmation requirement |
| Display | Nested hierarchy view | `list_live_notebooks` renders section groups as an indented tree instead of a flat section list |
| Fixed | Page-title escaping | Titles containing `]]>` could break the CDATA wrapper and corrupt a page's XML; titles are now escaped like bodies already were |
| Compatibility | Backup-folder detection | The local backup folder is auto-detected regardless of Windows/Office display language, instead of assuming the English name "Backup" |
| Distribution | `.mcpb` packaging | Packaged as a self-contained bundle for one-click installation in Claude Desktop |

See [CHANGELOG.md](CHANGELOG.md) for the version-by-version history.

## Prerequisites

- Windows, with the Microsoft OneNote desktop app installed
- For the backup-based reading tools: OneNote must have run at least once so a local backup exists
- For the live tools I made: OneNote just needs to be installed — its COM API starts it automatically if it isn't already running

## Installation

Drag the `.mcpb` file onto Claude Desktop (or double-click it) and confirm the install prompt. Claude Desktop manages the Python/`uv` runtime itself, so nothing else needs to be installed. Restart Claude Desktop afterwards.

## How it works

- The backup-reading tools parse the `.one` files OneNote automatically keeps under `%LOCALAPPDATA%\Microsoft\OneNote\16.0\<Backup>\`. The folder name is auto-detected regardless of Windows language (English "Backup", German "Sicherung", etc.) — it scans for whichever subfolder actually contains `.one` files.
- The live tools drive the OneNote desktop app through its COM automation interface (`OneNote.Application`), invoked via PowerShell under the hood.

## Credits & License

Originally created by [mhzarem](https://github.com/mhzarem/onenote-mcp) and released under the MIT license. I am not the original author — this is a modified, substantially extended derivative of that project. It stays open source under the same MIT license.
