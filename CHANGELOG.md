# Changelog

All notable changes to this project are documented here.

## [1.3] - 2026-09-12

This batch wore me out more than I expected. What started as "let me add a find-and-replace tool" turned into a full day of chasing undo/redo edge cases across every write tool in here. `insert_block_after` in particular fought me through two separate failed attempts before I found the actual root cause — turns out OneNote's own stored line-break format doesn't match what this server writes itself, so my first two fixes both looked right and both weren't. Good to finally have it all working and tested. I think it's in a good shape now that satisfies me.

### Added
- `find_and_replace_in_page` — find and replace exact text within a page
- `replace_last_block` — replace the last content block on a page, without needing to know or repeat what it currently says; requires `confirm=true` if the page has only one block, since that means replacing the entire body
- `insert_block_after` — insert a new paragraph immediately after existing text elsewhere on the page, not just at the end
- `undo_last_action` / `redo_last_action` — undo the most recent logged action, or redo the most recently undone one, backed by a persistent history log (`history.md`) that survives restarts
- `list_recent_actions` / `get_action_detail` — see what's changed recently, or the full detail behind one specific change
- `get_history_file_path` — look up where the action history log actually lives on disk
- `list_drop_folder_images` — list image files available in the default image drop folder

### Fixed
- Various debugging

## [1.2] - 2026-09-09

### Added
- OCR text extraction — `read_live_page`/`read_live_section` now also surface text OneNote automatically recognizes in embedded images (e.g. a stat block screenshot, a scanned PDF page), clearly labeled since OCR output can contain recognition errors
- `insert_image_from_file` — insert an image onto a page by local file path; only a path is ever passed to the tool, never the image bytes, keeping large images out of the conversation entirely
- A bare filename (e.g. `"Background.png"`) can be used instead of a full path if the image is in the default drop folder, which is resolved via Windows' actual configured Downloads location (not just the default `%USERPROFILE%\Downloads`) — this also works correctly if Downloads has been relocated to another drive

### Dependencies
- Added Pillow (image dimension/format detection)

## [1.1] - 2026-09-09

### Added
- `create_section_group` — create a new section group (folder) in a notebook
- `move_section` — move a section to a different section group, or back to the notebook's top level; also works to restore a section out of the recycle bin
- `rename_section` — rename an existing section
- `rename_page` — rename an existing page's title, leaving its content untouched
- `delete_section` / `delete_page` — delete a section or page (moves it to the OneNote recycle bin, not permanent). Requires an explicit second call with `confirm=true`; the first call only returns a preview and deletes nothing
- `list_recycle_bin` — list sections and pages currently in a notebook's recycle bin, to find what to restore or double-check what a delete removed

### Changed
- `list_live_notebooks` now renders the full nested hierarchy (section groups as an indented tree) instead of a flat list of sections

## [1.0] - 2026-09-08

First public release. Fork of [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp).

### Fixed
- Page titles containing `]]>` could break the CDATA wrapper and corrupt a page's XML — the original escaped this in page bodies but not in titles. Titles are now escaped the same way.

### Added
- `create_section` — create a new section in a notebook (the original could only create pages inside sections that already existed)
- `read_live_page` / `read_live_section` — read a single page or a whole section directly from the running OneNote app, instead of only from the (potentially outdated) local backup snapshot
- Language-independent backup-folder detection — the original hardcoded the English folder name `Backup`; it's now auto-detected regardless of Windows/Office display language
- Packaged as a `.mcpb` bundle for one-click installation in Claude Desktop
