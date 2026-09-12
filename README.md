I rebuilt this local MCP server so Claude can read and write my local Microsoft OneNote notebooks directly — no Microsoft account, API key, or Azure app registration needed, everything runs on my own machine. Windows only, since it drives the OneNote desktop app directly — no macOS/Linux support. It reads from OneNote's local backup snapshots for quick browsing and search, and talks directly to the running OneNote desktop app via its COM automation interface for always-current reads and for writing and organizing content.

Have fun with it. Between live edits, undo/redo, and no longer hunting through nested sections by hand, I've caught myself opening OneNote's actual window less and less — Claude just does it for me now, correctly, most of the time. If it saves you even half the clicking it's saved me, consider your evening reclaimed.

## Story

I run a D&D campaign, and years of session notes, NPC write-ups, quizzes, and my DM diary all live in OneNote. I didn't want to migrate all of that to Notion or one of the other note apps that already have an official Claude connector — OneNote is simply where the notes already were, and restructuring years of notebooks just to get a connector felt like the wrong trade-off. But I also didn't want to keep manually copy-pasting pages into Claude every time I needed context — "wait, what actually happened in session 4?" shouldn't mean digging through OneNote by hand every time. So instead of switching tools, I built the bridge myself.

This started as [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp), a small MCP server for reading local OneNote backup files and writing to OneNote via COM. It covered the basics, but couldn't read a single page or a whole section straight from the live app, couldn't organize sections at all, and had a small escaping bug in page titles. I picked it up and substantially extended it to close those gaps. It stays open source.

## Added Features

| Type | Feature | Description |
|---|---|---|
| Reading | `read_live_page` | Read the full text of a single page directly from the running OneNote app |
| Reading | `read_live_section` | Read every page in a section directly from the running app |
| Reading | `list_recycle_bin` | List sections and pages currently in a notebook's recycle bin |
| Reading | OCR text extraction | `read_live_page`/`read_live_section` also surface text OneNote recognizes in embedded images (e.g. a stat block screenshot), clearly labeled since OCR can contain errors |
| Writing | `insert_image_from_file` | Insert an image onto a page from a local file path — or by filename alone if it's in the default Downloads folder |
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
| Compatibility | Real Downloads-folder detection | Resolves Windows' actual configured Downloads location (via `SHGetKnownFolderPath`), so filename-only image lookups work even if Downloads was relocated to another drive |
| Distribution | `.mcpb` packaging | Packaged as a self-contained bundle for one-click installation in Claude Desktop |
| Writing | `find_and_replace_in_page` | Find and replace exact text within a page |
| Writing | `replace_last_block` | Replace the last content block on a page, without needing to know what it said — requires confirmation if the page has only one block |
| Writing | `insert_block_after` | Insert a new paragraph right after existing text elsewhere on the page, not just at the end |
| Undo/Redo | `undo_last_action` / `redo_last_action` | Undo the most recent change, or redo the most recently undone one — backed by a persistent history log that survives restarts |
| Undo/Redo | `list_recent_actions` / `get_action_detail` | See recent changes, or the full detail behind one of them |
| Undo/Redo | `get_history_file_path` | Look up where the action history log lives on disk |
| Reading | `list_drop_folder_images` | List image files available in the default image drop folder |

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

Originally created by [mhzarem](https://github.com/mhzarem/onenote-mcp) under the MIT license, and that's a real starting point I'm glad I didn't have to build from scratch, not a formality. But to be straight about where things stand now: the original could only read from local backup files and had a handful of basic write calls — no live reads, no organizing sections or pages, no editing, no undo. Everything that makes this a tool worth actually reaching for day to day — live reads and writes, section/page organizing, image insertion, find-and-replace, undo/redo, and more — I built. At this point my own additions outnumber what the original project ever had by a wide margin. Still stays open source under the same MIT license, and I'm not the original author of the concept or the backup-parsing groundwork it started from — just everything built on top of it since.
