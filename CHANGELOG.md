# Changelog

All notable changes to Slide Tools are listed here, newest first. Version numbers follow [Semantic Versioning](https://semver.org/): a major number for changes to how the panel is organised, a minor number for new features and a patch number for fixes.

## [2.1.0] 2026-11-12

### Changed

1. Resize: Resize from (the anchor picker) now shows at the More level as well as Full.
2. Resize: the formula help is now a small i icon at the top right of the tool. Hover over it, or tab to it, to read the help. It's available at every level.

## [2.0.0] 2026-11-12

### Changed

1. The panel is now a stack of dockable tool panels, one per tool, in the style of Illustrator's panels.
2. Grid, Layout presets, Position on slide and the preview are combined into a single Layout tool.
3. The preview sits inside the Layout tool instead of staying pinned over everything else.
4. Draw on slide and Fit selection moved from the bottom bar into the Layout tool.
5. Max width and Max height are now called Area width and Area height.
6. The code is split into separate files per tool so new tools can be added without touching the others.

### Added

Panel:

1. Each tool has detail levels (Basic, More, Full), cycled with a button in its header. Tools only show the levels they need.
2. Fold any tool down to its header.
3. Drag tools by the grip to reorder them, or use Move up and Move down in the tool menu.
4. Hide and show tools from the settings button, plus Unfold all, Fold all and Reset panel arrangement.
5. Reset any tool to its defaults from its menu.
6. Order, detail levels, folding and hidden tools are remembered.
7. Dark theme that follows PowerPoint.

Layout:

1. Use selection as area sets the grid position and size from the selected shapes.
2. Save your own presets with a name, and delete them again.
3. New shapes are selected after drawing, and can optionally be grouped.
4. Selected shapes appear as dashed outlines on the preview.

Resize:

1. Lock aspect ratio.
2. Resize from any of nine anchor points.
3. Match selected shapes to the largest or smallest.
4. Recent formulas, with one click to reuse.

Corners:

1. Radius presets (2, 4, 8, 12 and 16 pt).
2. Read radius from a selected rounded rectangle.
3. Press Enter in the Radius box to round the selection.

Project:

1. README with install, usage, troubleshooting and a guide to adding tools.
2. This changelog.
3. `slidetools-manifest.xml` kept in the repository, with the corrected `Slidetools` address.

## [1.0.0] 2026-11-12

### Added

1. First release of the side panel: grid calculator with presets and preview, Draw on slide, Fit selection, formula based Resize, and Corners for rounded rectangles.
2. Settings saved between sessions.
