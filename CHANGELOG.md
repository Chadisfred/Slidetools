# Changelog

All notable changes to Slide Tools are listed here, newest first. Version numbers follow [Semantic Versioning](https://semver.org/): a major number for changes to how the panel is organised, a minor number for new features and a patch number for fixes.

## [3.11.0] 2026-11-26

### Added

1. Start with this presentation, in the settings menu: Slide Tools loads by itself when the file opens, so shortcuts and follow selection work at once. Show the panel when it starts opens the panel with the file too. Both are saved in the file, so a company template carries them to every deck made from it.
2. Keyboard shortcut Ctrl+Alt+Shift+P shows the Slide Tools panel.

### Changed

1. Elements opens with a picture of each element, all the same size and drawn in the chosen colour. Picking one shows its settings.
2. Elements: white text and white bands are always white, even when a template's Light 1 theme colour is grey or mapped to a dark colour.
3. The Home tab button is now called Slide Tools Panel, in a Slide Tools group, instead of Grid Panel in Slide Panel. Run the installer once more to get the new name and the new shortcut.

### Fixed

1. Elements: Type a size places the element when nothing is selected. It used to ask for a selection first.

## [3.10.0] 2026-11-26

### Added

1. Elements, a new tool for prebuilt design elements in theme colours and fonts: Stat card (a bordered card with a header holding a title and a value with an automatic up, down or flat arrow, a summary band, bullets and a footer figure, with rules between the bands in the header colour, as thick as the outline), Pointer card (a dark header, a band of bullets and a light footer joined by small pointers) and Big number. An element replaces the selected shape at its size and place, or takes a typed width such as `3col` at the grid area's top left.

## [3.9.0] 2026-11-26

### Changed

1. Animation is laid out like an inspector: Entrance, Emphasis or Exit first, then a searchable gallery of effects with category chips, the preview, and folding Effect, Motion, Timing and Presets sections that show a summary when closed. Copy, Paste, Repeat and Remove sit in one row, and Apply to selection stays at the bottom while you scroll.

## [3.8.0] 2026-11-25

### Added

1. Animation: Repeat for emphasis. Play a pulse, spin, dim or Attention effect such as Flash a number of times, until the next click, or until the end of the slide.
2. Animation: Copy animations and Paste animations. Copy the selected shapes' animations as a set, keeping their order, clicks, with and after, and delays, and paste the set onto other shapes on any slide.
3. Animation: Repeat set adds the selected shapes' animations again, after themselves.

## [3.7.1] 2026-11-25

### Fixed

1. Animation, Images and Copies find the selected shape more reliably in the slide: by its box, then by its name when its position reads slightly differently, then by its id. Shapes that use newer PowerPoint features with a fallback are found too.
2. When a shape still can't be found, the message names it and says why, for example when it sits inside a group and the whole group needs selecting.

## [3.7.0] 2026-11-25

### Added

1. Tables: Row heights, alongside Column widths. Pick rows with chips that show their height in g, then type a height such as `g`, `g/2`, `+g` or `*1.5`. Round to g and Equal work on rows too. A row whose text needs more room stays taller, and the panel says which rows did.

## [3.6.0] 2026-11-24

### Added

1. Tables: Column widths. Pick one or more columns with chips that show their width in g, then type a width such as `2g`, `col`, `+g` or `*1.5`. Round to g and Equal adjust the chosen columns too. The change is direct, with no slide swap.
2. Animation: 44 keyframed effects in new groups, Fade and slide, Back, Bounce, Zoom, Roll and spin, Attention and Exits, including Bounce in, Jack in the box, Rubber band, Heartbeat, Tada, Wobble and Back out. Several are inspired by Animate.css (MIT).

### Changed

1. Animation: presets sit in groups that fold open one at a time, so the list stays short.
2. Animation: keyframed effects use GSAP's elastic, back, expo and power curves, so overshoots and slides move smoothly instead of in stops and starts. Curves are written with about one point every 20 ms, which adds around 1 KB per effect to a saved file.

## [3.5.0] 2026-11-24

### Changed

1. Typography keeps only the settings that change the slide directly, so every change shows on the slide straight away: font, size, style, case, colour, margins, autofit and wrap.
2. Position now includes Snap, the same as in Grid. The two share their settings.

### Removed

1. Typography: character spacing, line spacing, space before and after, kerning and baseline. Add-ins can't set these directly, so they needed a slide swap and couldn't be seen before applying. PowerPoint's own Paragraph and Font settings cover them.
2. Typography: the live preview. It couldn't use Office cloud fonts, so it didn't match the slide.

## [3.4.2] 2026-11-24

### Fixed

1. Images: pasting a new picture onto an existing one no longer keeps the old image in the file unseen. Any slide change made by the panel now drops picture links the slide no longer uses, which keeps files smaller.

## [3.4.1] 2026-11-24

### Changed

1. Size: steps of g show as decimals, such as 2.5g or 3.375g, which are easier to read than fractions.

## [3.4.0] 2026-11-24

### Added

1. Size: the line under Width and Height shows sizes in steps of g as well as inches, such as 6g x 2½g. With the boxes empty, it shows the selected shape's own size. Sizes on whole, half, quarter or eighth steps show as fractions, and others as about, for example about 3.33g.

## [3.3.0] 2026-11-24

### Added

1. Animation, a new tool. Build entrances, exits and emphasis that PowerPoint's menus can't make: real easing curves with GSAP's names (power, sine, expo, circ, back, elastic, bounce, with in, out and in out), several effects at once (fade, move, scale and rotate), exact durations, delays and distances (including grid names), and a preview in the panel that plays the same keyframes written to the slide.
2. Animation: a stagger for several shapes, in picked, left to right or top to bottom order.
3. Animation: starter presets (Fade up, Pop, Slide in, Drop in, Zoom in, Swing in, Elastic grow, Rise out, Shrink out, Pulse, Gentle spin), your own presets saved in the presentation, Capture to reuse an existing shape's animations, and Remove animations from selection.

## [3.2.1] 2026-11-24

### Fixed

1. Startup is more robust. The panel tells PowerPoint it is ready straight away instead of after building every tool, its scripts load without holding up the page, and the update check waits until after launch. This makes the occasional "Error loading add-ins" message at launch less likely.
2. A tool that fails to start shows a note in its card instead of stopping the rest of the panel and the keyboard shortcuts.

## [3.2.0] 2026-11-23

### Added

1. Typography: minus and plus on Character spacing, Line spacing, Space before and Space after. Each box shows the value in use now, read from the shape or its layout, master or default text style, and the steps start from it. The preview shows the result straight away, and Apply spacing sets it.

### Fixed

1. Typography: the Font box no longer turns red for a font name.
2. Typography: the preview shrinks the text the same way PowerPoint does when Shrink text on overflow is on, so it fits the dashed box as it does on the slide.
3. Typography: the preview tries a font's family at the right weight, such as Avenir Next LT Pro at Light, and says when a font isn't available to the panel, as with Office cloud fonts, instead of quietly showing another font.
4. Typography: Case buttons no longer look half on when PowerPoint can't report them.

## [3.1.1] 2026-11-23

### Fixed

1. Typography: selecting a shape with no text no longer shows a large empty preview. The preview only appears when there is text to show.

## [3.1.0] 2026-11-23

### Added

1. Typography: font controls. Font (with the theme's heading and body fonts first in the list), Size with minus and plus, Bold, Italic and Underline, and at More, All caps, Small caps and Strike. They apply straight away to the highlighted text, or to all the text in the selected shapes, and show the current settings.
2. Typography: a text colour button that opens Colours set to Text, with Typography pinned beside it.
3. Typography: every points box takes grid formulas in points, such as `2g` for a size, `g/2` for a margin or `g` for line spacing. Size also takes `+2` or `*1.5`, worked out from each text's own size. A box that can't be read turns red and says why.

## [3.0.0] 2026-11-23

### Changed

1. Tool buttons at the top of the panel replace the long scrolling stack. Every tool is named in one small grid, grouped by job: Grid, Position, Size and Copies, then Typography, Colours, Text styles and Shape styles, then Images, Tables, Corners and Similar. Clicking a button opens that tool and closes the others, so there is little to scroll. Show all tools in one list, in the settings menu, brings back the stacked view.
2. Arrange, Transform, Layout and Match are regrouped by job. Layout is now Grid and holds Snap, Arrange is now Position and holds copy and paste position, Transform is now Size, and Match is now Select similar. In detail:
   1. Layout is now Grid, and Snap (guides, columns, steps and the guide report) moves into it. Use in Layout is now Use for grid.
   2. Arrange is now Position, and copy and paste position and size moves into it from Match.
   3. Transform is now Size: width and height, lock, reference point and match size.
   4. Match is now Select similar.
3. Saved panel arrangements start once from the new grouped order. Settings, styles, presets and keyboard shortcuts carry over unchanged.

### Added

1. Copies, its own tool, split out of Transform with all the same settings. It shares the reference point with Size.
2. Pin a tool from its header to keep it open while you open another, such as Size and Position together.
3. The panel follows the selection: selecting a picture opens Images, a table opens Tables, and clicking into text opens Typography. Selecting shapes never changes what is open. Open tools with the selection, in the settings menu, turns this off.

## [2.31.0] 2026-11-23

### Added

1. Typography: a live preview. A scaled copy of the first selected text box, with its font, size, colour, width and margins, updates as you type spacing and margin values, so you can see line breaks and tightness before applying. The dashed line marks the shape's height. A switch at More hides it.

### Fixed

1. Typography: the Autofit buttons are now Off, Shrink and Resize, and margin labels are Left, Right, Top and Bottom, so nothing breaks onto two lines on narrow panels.
2. On panels 330 px wide or less, the level button shows just its dots, so tool titles aren't cut off, and the fold button no longer sticks out past the edge.

## [2.30.0] 2026-11-23

### Added

1. Typography, a new tool. Set exact character spacing, line spacing in points or lines, space before and after paragraphs, kerning and baseline for all the text in the selected shapes. Empty boxes are left alone. Margins, autofit and wrapping apply straight away. Read from selection copies a shape's settings into the boxes.

## [2.29.0] 2026-11-23

### Added

1. Images, a new tool. Paste into shape turns the selected shapes into real pictures with the same position, size, corners and outline, cropped to fill without stretching. The whole picture is kept, so PowerPoint's Crop button can pan and zoom it afterwards. Pasting onto a picture swaps its image, and an empty picture placeholder stays a placeholder.
2. Images: Crop position moves the crop of selected pictures to any of nine points, keeping how much is cropped, and sets where new pictures are cropped from. Reset crop shows the whole picture again.

### Changed

1. Crop pictures to fit moves from Transform to Images.

### Removed

1. Transform: Picture fill, replaced by Images, since a filled shape's crop couldn't be adjusted afterwards.

## [2.28.0] 2026-11-23

### Added

1. Versions and updates in the settings menu. The top of the menu shows the panel and install versions. The panel checks for updates when it opens and puts a dot on the settings button when one is ready. Update panel reloads it with the new version, and a short What's new list shows the changes. When a new install file is out, Download installer gets it, ready to double click.

### Changed

1. The install file tells the panel which install it is. Run the installer once more so the panel can track it.

## [2.27.2] 2026-11-23

### Changed

1. Size shortcuts: Ctrl+Alt+Shift+I now makes the selection taller and K makes it shorter.

## [2.27.1] 2026-11-23

### Fixed

1. Changed shortcut keys could be ignored because PowerPoint kept an old copy of the key list. The install file now gives the list a versioned address, so PowerPoint fetches the current keys. Run the installer once more to pick this up.

## [2.27.0] 2026-11-23

### Added

1. Transform: Picture fill. Select shapes, then paste, drop or choose a picture to fill them like a picture placeholder. Fill and crop crops the picture to each shape's proportions so it never stretches, with the crop following the reference point. Fit inside, at More, shows the whole picture. The picture is cropped inside the panel and stays on this computer. Paste again to swap it.

## [2.26.1] 2026-11-23

### Changed

1. The size shortcuts move to Ctrl+Alt+Shift+I, J, K and L, laid out like the arrow keys (L wider, J narrower, K taller, I shorter, Cmd+Option+Shift on a Mac), because PowerPoint keeps Alt+Shift with the arrows for moving text.

### Removed

1. The align and distribute shortcuts, to free the keys. The Align and Distribute buttons in Arrange are unchanged.

## [2.26.0] 2026-11-23

### Added

1. Size shortcuts: Alt+Shift+Right and Down make the selection wider and taller, and Left and Up make it narrower and shorter, by Arrange's Nudge by amount (Option+Shift on a Mac). Sizes change around Transform's reference point, a straight line only changes along its length, and nothing shrinks below 1 pt.

## [2.25.1] 2026-11-23

### Fixed

1. Keyboard shortcuts could show "This add-in could not be started". The install file now names the panel page as the add-in's function file, so PowerPoint can start Slide Tools when a shortcut is pressed, and the shortcut actions are registered as soon as the page loads. Run the installer once more to pick this up.

## [2.25.0] 2026-11-23

### Added

1. Keyboard shortcuts that work while the slide has focus: Ctrl+Alt+Shift with an arrow to nudge, D to repeat the last action, L, C, R, T, M and B to align, and H and V to distribute (Cmd+Option+Shift on a Mac). They need PowerPoint for Windows 2601 or later, PowerPoint for Mac 16.105.2 or later, or PowerPoint on the web. Keyboard shortcuts in the settings menu lists them, and the Arrange buttons show their keys in their tooltips.

### Changed

1. The install file now uses a shared runtime and links the shortcut keys. Installs from before 2.25.0 need the installer run once more for the shortcuts. Everything else keeps working without it.

## [2.24.0] 2026-11-23

### Changed

1. Arrange: Nudge is now part of the Basic level.
2. Transform: when a Move box can't be read, the box turns red and the message says what to change, such as "Move down ends with +", or which name isn't known. The message shows at every level, before Make copies is clicked.

### Removed

1. Colours: the Contrast check.

### Fixed

1. Colours: colouring text could fail with "PowerPoint didn't accept one of the values". Text without a Slide Tools style is now coloured in one step, and styled text falls back to a plain colour if PowerPoint refuses the detailed edit.

## [2.23.0] 2026-11-23

### Added

1. Colours: a Contrast check at the More level. For the selected shape it shows the text colour on the fill, the contrast ratio, and whether it passes for body text (4.5 to 1) and large text (3 to 1). When body text fails, it suggests a passing colour from the theme, your colours, white or black, and applies it with one click while keeping style links.
2. Tables: Paste data. Cells copied from Excel build the new table, with the first row as the header and column roles worked out from the values. Quoted cells with line breaks are handled, up to 40 rows and 20 columns.
3. Tables: Line up decimals, a style setting on by default. Numbers in data columns are padded to the same number of decimal places so the points line up, keeping currency signs, separators, percent signs and brackets. Only cells that change are rewritten.

## [2.22.0] 2026-11-23

### Added

1. Layout: Column ratio and Row ratio for unequal columns and rows, such as `2:1:1`. The number of parts sets Columns or Rows, Draw, Fit and the preview follow the ratios, and Arrange's Columns snapping uses the same cells. Ratios are saved in presets.
2. Transform: Rows for copies, which makes a grid in one step. Each row repeats the original and its copies across, one Move down lower, or the shape's height plus the row gap apart when Move down is empty.

### Changed

1. Layout: the cell size readout lists each width and height when they differ, in a smaller size when it is long.

## [2.21.0] 2026-11-23

### Added

1. Share settings, in the settings menu. Export copies a settings code with any of your text styles, shape styles, table styles, colours and Layout presets. Import takes a pasted code and merges it (adding new items and updating ones with the same name) or replaces the chosen lists. Shape styles stay linked to their text styles.
2. A Repeat button in the top bar runs the last action again on the current selection with the same values: Resize, Match size, Make copies, Align, Distribute, Even gaps, Snap, Round corners, and applying a style or colour.
3. A Replace this slide? message before cropping pictures or making copies, explaining that comments may not carry over and Undo takes more than one step, with Don't ask again. Cancelling a crop keeps the pictures in proportion instead. Ask before replacing a slide again, in the settings menu, brings the message back.

## [2.20.2] 2026-11-23

### Changed

1. The install guide is now a page in the manual, in the same design, linked from the top bar, the sidebar and Getting started. INSTALL.md is a short pointer to it.

## [2.20.1] 2026-11-23

### Changed

1. Manual: the same design polish as the panel. Content, layout and navigation are unchanged.
   1. Sidebar and table headings are in normal case instead of small spaced capitals, and headings are a little lighter.
   2. Numbered markers and table numbers are ink circles with monospaced digits, and step numbers are outlined circles.
   3. Level tags show the panel's dots: one for Basic, two for More, three for Full.
   4. Diagram guides are neutral dashed lines, code and keys use the panel's monospaced face, and spacing sits on a 4 px grid.
   5. Markers show a ring on hover instead of growing.

## [2.20.0] 2026-11-23

### Changed

1. A design polish across the panel. Every tool, control and behaviour is unchanged.
   1. Section labels are in normal case instead of small spaced capitals, and the Layout tiles read Across, Down, Grid and Cells.
   2. Measurements (cell size, selection details, tile values and result lines) use a monospaced face with even digits, in ink instead of blue.
   3. Blue is kept for main buttons, selections and switches. Tool icons and stepper buttons are grey.
   4. The Layout preview shows the grid area as a slide with a dashed guide on a grey canvas, with flat cells and ink outlines for selected shapes. The Copies preview uses the same canvas.
   5. Sizes sit on a 4 px grid, buttons match the height of input boxes, badges and chips have small rounded corners, and nothing is heavier than weight 600.
   6. Motion uses short named fades with one easing curve. Swatches show a ring on hover instead of growing.
   7. Single sentence status lines no longer end with a period.

## [2.19.0] 2026-11-20

### Added

1. Transform: Copies, working like Illustrator's Transform effect. Set a number of copies, Move across and down, Scale across and down, Rotate and Reflect X and Y, around the reference point. Each copy is transformed from the one before. Space by Distance or Gap. All boxes take formulas such as `w+gap`, and a live preview shows the result. Copies are exact, including text, formatting, pictures, groups and tables, and are selected afterwards.

### Changed

1. Resize is now called Transform. Saved settings, order and levels carry over.
2. Resize from is now Reference point, shared by resizing, matching, cropping and copies.
3. Picture cropping and copies share one in-memory slide rebuild. Only the slide is edited, never masters or layouts, and nothing leaves the panel.

### Fixed

1. Layout: the gap rows show all six choices on one line.

## [2.18.0] 2026-11-17

### Added

1. Layout: None in both gap rows, so cells sit edge to edge.
2. Resize: lines are resized by their length from either box, keeping their direction. Horizontal, vertical and diagonal lines all work, and the result line shows the length.
3. Resize: Crop pictures to fit. When a picture's frame changes shape, the picture is cropped to fill it instead of being stretched. The picture keeps its scale and the Resize from point stays still, and it grows only when the frame would run past its edges. The crop is written to a copy of the slide in memory, which replaces the original. Masters and layouts are never edited, and nothing leaves the panel.

### Changed

1. Resize: Match size compares lines by length and matches them to each other, separately from other shapes.
2. Resize: Use selection puts a line's length in Width.

## [2.17.0] 2026-11-17

### Added

1. Tables: spaces, tabs and line breaks at the start and end of each cell are trimmed before formatting, so text pasted from Excel wraps properly and column widths fit the real text. Spaces inside the text are kept, and only cells that change are rewritten. A style setting turns it off.
2. Tables: a Trim spaces button for a one-off clean up of the selected table.

## [2.16.0] 2026-11-17

### Added

1. A full manual page for every tool in `manual/`: overview and quick start, an annotated screenshot with a numbered marker on each control, a reference table covering every button, box, switch and chip (level, what it does, values and defaults, notes), how it works, recipes, limits and edge cases, messages and what to do, and related tools.
2. Reference pages: Formulas and units, Glossary, and Expert tips.
3. A new manual home page with panel basics, what is saved where, PowerPoint requirements and troubleshooting.
4. Shared navigation on every page: top bar, a sidebar of all pages, an "On this page" list, and Previous and Next links. Light and dark themes, and readable on a phone.

### Changed

1. How to use in each tool's menu now opens that tool's own manual page.

## [2.15.0] 2026-11-17

### Added

1. User manual at `manual.html`: getting started, then a section per tool with quick start steps, a before and after example, panel screenshots, what each level shows, and tips. Light and dark, and readable on a phone.
2. How to use in every tool's ... menu, opening that tool's section of the manual.

### Changed

1. Every tool's i tip is now one line saying what the tool does, followed by up to three steps. Layout and Corners have one for the first time.
2. Explanations that sat under controls moved into small i icons next to the setting they describe (grid spacing, gaps, formulas, corners, distribute, snap modes, match by, paste from, table columns, table fonts, column fill and sizes).
3. Shorter empty states and clearer button hover text.

## [2.14.0] 2026-11-16

### Changed

1. Tables: with only one category column, the lines between groups are 1/4 pt instead of 1 pt. The header, total and bottom rules stay 1 pt. A setting in the style editor turns this off.

## [2.13.2] 2026-11-16

### Fixed

1. Tables: tables inside content placeholders, which is how the table icon in a layout creates them, weren't recognised, so applying a style said no table was selected. They're now found when selected, with the cursor in a cell, and by Update all.
2. Tables: when the selection isn't a table, the message names what is selected.
3. Tables: total rows are ignored when guessing column roles, so an empty label cell in a total row no longer turns a label column into a category.

### Added

1. Tables: a line at the More level showing what PowerPoint reports about the selection.

## [2.13.1] 2026-11-16

### Fixed

1. Tables: clicking a table style could say "Select a table first" even with a table selected, most often on the first click when the line check runs. The table is now found before the check and selected again afterwards.
2. Tables: with the cursor inside a cell, where PowerPoint reports no selected shape, the panel uses the table it last saw on the slide, or the only table on the slide. The column chips show in that case too.

## [2.13.0] 2026-11-16

### Added

1. Tables: grid-based sizes. The table fills the Layout grid area in whole steps of the grid spacing g, category and data columns fit their text, label columns take the rest, and rows are 1g tall, growing in 1/4 g steps when text wraps. Adjustable or switchable in the style editor.
2. Tables: Check lines, a one-time test that finds out how this PowerPoint shows table lines, remembered on this computer and repeatable from the More level.

### Fixed

1. Tables: on PowerPoint versions that apply table line transparency the opposite way round, lines meant to show were hidden and lines meant to be hidden showed as thin grey lines. Lines now follow the result of the line check.

## [2.12.1] 2026-11-16

### Fixed

1. Tables: Insert table could fail with "PowerPoint didn't accept one of the values". It now adds a plain table first and then formats it, falling back to a simpler request if PowerPoint refuses the first one.
2. Tables: formatting is sent in separate groups (lines, fill, fonts, padding, alignment). If PowerPoint refuses one group, the rest still apply, and the message names the group that was refused.

## [2.12.0] 2026-11-16

### Added

1. Tables tool: named table styles saved in the presentation. Insert a formatted table at the Layout grid area, or format the selected table with one click.
2. Column roles per table (Category, Label, Data), guessed from the contents and changed with a click. Roles set alignment, which columns get group rules, and where the thin vertical lines go.
3. Default table style: 1 pt rules under the header and the table, 1 pt rules above new groups and total rows, 1/4 pt row lines, banded 2nd, 4th, 6th... columns, Avenir Next LT Pro fonts, and 0.05 in padding with flush outer edges.
4. Style editor for fonts (or text styles), every line's weight, colour and visibility, column fill, padding and the total row word. Update all reformats every linked table.
5. Stack repeats: blanks category labels that repeat the row above within the same group.

## [2.11.0] 2026-11-16

### Added

1. INSTALL.md: a step by step install guide for new users, with quick install, manual install, uninstall and troubleshooting.
2. A one click installer, `install/Install-SlideTools.cmd`, that downloads the install file into the user's folder and adds the per user PowerPoint setting. No admin rights needed.
3. A matching uninstaller, `install/Uninstall-SlideTools.cmd`.

## [2.10.1] 2026-11-13

### Fixed

1. Arrange: Guides found no guides in decks whose guides are presentation guides (the kind you can drag on any slide, usually brought in by the template). The panel now reads the open file, read only, and finds presentation guides as well as the current slide's layout and master guides.
2. Arrange: the guide report says where the guides were found, and what was checked when none were.

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
4. The install file's provider name is now "Slide Tools". Manifest version 1.0.0.2.

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
