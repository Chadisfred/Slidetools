# Changelog

All notable changes to Slide Tools are listed here, newest first. Version numbers follow [Semantic Versioning](https://semver.org/): a major number for changes to how the panel is organised, a minor number for new features and a patch number for fixes.

## [2.10.0] 2026-11-13

### Added

1. Arrange: Snap to template guides. The guides built into your template are read from the current slide's layout and master, and shapes snap to the nearest real guide. Snap all four edges works with guides too.
2. Arrange: a report on the template's guides at the More level, with their spacing and any guide slightly off the pattern, and Use in Layout to copy the spacing and starting position into the Layout tool.
3. Number boxes accept units: `30pt`, `1.5cm` and `12mm` are converted to inches.

### Changed

1. Arrange: the snap modes are now Guides, Columns and Steps (previously Grid cells and Fine grid), shown at every level, with a line describing what the chosen mode does.
2. The snap button is now just Snap, and the Grid heading is now Snap.

### Fixed

1. Number boxes rounded values to three decimal places when you clicked away, so a precise spacing such as 5/12 became 0.417 and shapes drifted from 30 pt guides across the slide. Values now keep five decimal places.
2. Use selection as area keeps five decimal places too.

## [2.9.0] 2026-11-13

### Added

1. Match tool: select every shape on the slide that matches the selected one by fill colour, outline, shape style, text style, size or type.
2. Match: Previous slide and Next slide jump to the next slide with matches and select them there, with a count for the slide and the whole presentation.
3. Match: copy a shape's position and size and paste the position, the size or both onto selected shapes, lined up by any of nine anchor points.

### Changed

1. PowerPoint errors are shown as plain sentences, for example "A shape or slide changed while this was running", instead of PowerPoint's own wording.
2. Styles: scanning the slides for Update all and font suggestions takes about a third as many trips to PowerPoint, which is noticeably faster on large presentations.
3. Colours and Layout load only the slide details they need, which helps on large presentations.

### Fixed

1. Resize with Lock aspect ratio on a straight line no longer produces an invalid size.
2. Draw on slide and Fit selection explain why when the grid doesn't fit its area, instead of doing nothing.

## [2.8.0] 2026-11-12

### Changed

1. Colours: Theme, Template and Yours are now tabs under the Fill, Outline and Text switch, and the panel remembers the last one used.
2. Colours: template colours are split into single colours and colour families (a colour with its tints in a column). Empty white slots are left out, and the grids keep their columns at any width instead of wrapping.
3. Narrow panels hide the detail level word in tool headers so the fold arrow stays on screen.
4. The install file's provider name is now "Slide Tools" instead of a personal name. Manifest version 1.0.0.2.

## [2.7.0] 2026-11-12

### Added

1. Colours: the five rows of tints and shades under each theme colour, worked out the same way as PowerPoint's colour menu and named the same way, for example "Accent 1, Lighter 40%".
2. Colours: Template colours, the Custom Colors saved in the template's theme, read from a copy of the first slide without changing the deck.

### Changed

1. Colours: theme colours now follow PowerPoint's order and names (Background 1, Text 1, Background 2, Text 2, Accent 1 to 6).

### Removed

1. The Consistency check tool. It was too granular to be useful, listing hundreds of tiny grid offsets on a single slide.

## [2.6.0] 2026-11-12

### Added

1. Consistency check tool: one scan of normal slides for shapes nearly on the grid, text with no style, unlinked text and shape lookalikes, odd corner radii, and colours outside the palette.
2. Click an item to select it on its slide. Snap, Link and Set fixes per item, and Fix all per group.
3. Choose which checks run and the grid tolerance at the More level.

### Changed

1. Each tool's i icon now sits in its header, so it no longer covers the tool's first row.

## [2.5.0] 2026-11-12

### Added

1. Colours tool: theme colours (read only) and your own palette saved in the presentation, applied to fill, outline or text with one click.
2. Add colours by hex, colour picker, or from the selected shape's fill or text's colour. Rename, reorder and remove them at the More level.
3. Recolouring styled text from the Colours tool keeps its text style link, recording the colour as a local change.
4. The colour pickers in the Styles and Shape styles editors suggest palette and theme colours.

## [2.4.0] 2026-11-12

### Added

1. Arrange tool: align left, centre, right, top, middle or bottom, relative to the selection, the Layout grid area or the largest shape.
2. Distribute across or down with an exact gap that accepts formulas such as g, and even gaps that keep the outer shapes still.
3. Snap to grid: to the nearest Layout cell corner, optionally resizing to whole cells, or to a fine grid in steps of g, optionally snapping all four edges.
4. Nudge by an amount or formula, and swap two shapes.

## [2.3.0] 2026-11-12

### Added

1. Shape styles tool, saved inside the presentation, covering fill (solid or none, transparency), outline (on or off, colour, weight, dash, line type, transparency), corner radius, text box margins, vertical alignment, wrap and autofit.
2. A shape style can include a text style, applied to the text inside.
3. Sample a shape, apply with one click, and Update all for shapes, which keeps hand changes on linked shapes and offers lookalikes with a choice to leave them, apply the style, or apply it and keep their colours.
4. Selection status at the More level with reset, unlink and link buttons.

## [2.2.1] 2026-11-12

### Fixed

1. The i info tips were cut off by the edge of their tool, so the Resize tool had to be at Full to read the whole tip. Tips now open over the rest of the panel, above the icon when there isn't room below, and scroll if they're taller than the panel.

## [2.2.0] 2026-11-12

### Added

1. Styles tool for character styles, saved inside the presentation.
2. Sample a style from selected text, or start from a blank style. Mixed settings are left out.
3. Apply a style to selected text or whole text boxes with one click.
4. Edit styles, then Update all pushes the edits to every normal slide after showing what will change. Masters and layouts are never touched.
5. Local changes on linked text are kept when a style is updated.
6. Lookalike detection for text that matches a style but has lost its link, for example after recolouring from the ribbon. Choose to leave it, apply the style and colour, or apply the style and keep its colour, for the whole group or item by item.
7. Selection status at the More level shows the style in use, a "+" for local changes, and buttons to reset, unlink or relink.
8. Font suggestions in the style editor: fonts used in the presentation, fonts in saved styles, and installed fonts where PowerPoint's browser allows listing them. A refused or unavailable font check is remembered for about 12 hours so you aren't asked again.
9. Weight safeguards: weighted fonts such as Semibold, Black or Light get Bold set to Off automatically, with a warning if you switch it back on. Applying a weighted font with Bold on Leave switches leftover bold off. Bold On with no font warns about fake bold.

### Changed

1. Reset to defaults is greyed out in the tool menu for tools that have nothing to reset.

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
