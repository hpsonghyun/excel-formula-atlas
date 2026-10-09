# Changelog

User-visible changes are recorded here. This repository is being prepared for its first public release.

## Unreleased

### Fixed

- Arrowheads now follow the source-to-formula direction and meet the incoming edge of the target cell. Short same-sheet connections no longer overshoot and turn back inside the cell. Contained ranges and self references keep a nonzero loop.

### Added

- A single pasteable VBA module that analyzes the active workbook or a file selected with the Windows file picker and exports an interactive HTML map.
- Worksheet maps that keep referenced cells and ranges at their captured row and column positions.
- Upstream and downstream tracing, selectable connections, area selection, search, and a nearby-table inspector.
- Right-button panning, mouse-wheel zoom, and keyboard controls for common map actions.
- Loading stages with completed item counts, a work-based percentage, elapsed time, and a visible error state.
- Chunked rendering for large worksheets and workbook maps.
- An English default interface with a Korean language option.
- A bundled local Windows helper for navigation from a connected map to Excel.
- A fictional operations workbook, a static interactive demo, and an English project page.

### Notes

- Generated maps contain workbook snapshots. They must be regenerated after the source changes.
- Reference analysis and worksheet capture are bounded; advanced and dynamic reference forms may not be fully resolved.
- Browser tests, helper component checks, and real Excel execution are separate validation scopes. Release notes must identify which scopes have been verified.
