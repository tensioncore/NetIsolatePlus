# Changelog

## [1.0.0] - 2026-01-17
### Added
- Adapter list with per-adapter toggle, Isolate, and Status actions
- Bulk “Toggle All” switch
- Sorting options (Enabled first / Disabled first / Name A→Z / Name Z→A)
- Optional filtering of virtual adapters
- Start with Windows toggle (Task Scheduler-based)

### Changed
- Adapter identity is GUID-only to ensure stability across renames
- Improved responsiveness with busy overlay during long operations
- Isolation mode protects UI from conflicting toggles

### Fixed
- Startup reliability on Windows 10/11 (GUI visible on login)
- Reduced UI glitches during bulk toggles and isolation workflows