# Changelog

All notable changes to this project are documented here.

## [1.1.0] - 2026-09-09

### Added
- `create_section_group` — create a new section group (folder) in a notebook
- `move_section` — move a section to a different section group, or back to the notebook's top level; also works to restore a section out of the recycle bin
- `rename_section` — rename an existing section
- `rename_page` — rename an existing page's title, leaving its content untouched
- `delete_section` / `delete_page` — delete a section or page (moves it to the OneNote recycle bin, not permanent). Requires an explicit second call with `confirm=true`; the first call only returns a preview and deletes nothing
- `list_recycle_bin` — list sections and pages currently in a notebook's recycle bin, to find what to restore or double-check what a delete removed

### Changed
- `list_live_notebooks` now renders the full nested hierarchy (section groups as an indented tree) instead of a flat list of sections

## [1.0.0] - 2026-09-08

First public release. Fork of [mhzarem/onenote-mcp](https://github.com/mhzarem/onenote-mcp).

### Fixed
- Page titles containing `]]>` could break the CDATA wrapper and corrupt a page's XML — the original escaped this in page bodies but not in titles. Titles are now escaped the same way.

### Added
- `create_section` — create a new section in a notebook (the original could only create pages inside sections that already existed)
- `read_live_page` / `read_live_section` — read a single page or a whole section directly from the running OneNote app, instead of only from the (potentially outdated) local backup snapshot
- Language-independent backup-folder detection — the original hardcoded the English folder name `Backup`; it's now auto-detected regardless of Windows/Office display language
- Packaged as a `.mcpb` bundle for one-click installation in Claude Desktop
