# Slide Tools

A PowerPoint side panel for laying out grids, resizing shapes with formulas, rounding corners, selecting and matching similar shapes, and managing character, shape and table styles. It runs as an Office web add-in and is hosted on GitHub Pages:

https://chadisfred.github.io/Slidetools/taskpane.html

Open that address in a browser to try the panel in preview mode. The calculator works there, and the slide buttons work once the panel is opened from PowerPoint.

Note the capital S in `Slidetools`. The address doesn't work with a lowercase s.

New to the tools? The [user manual](https://chadisfred.github.io/Slidetools/manual.html) has a full page for every tool: an annotated screenshot with a reference entry for every control, how it works, recipes for real tasks, limits and messages. It also has reference pages for [formulas and units](https://chadisfred.github.io/Slidetools/manual/formulas.html), a [glossary](https://chadisfred.github.io/Slidetools/manual/glossary.html) and [expert tips](https://chadisfred.github.io/Slidetools/manual/tips.html). In the panel, How to use in each tool's ... menu opens that tool's page.

## Contents

1. [Installing](#installing)
2. [Using the panel](#using-the-panel)
3. [The tools](#the-tools)
4. [Keyboard shortcuts](#keyboard-shortcuts)
5. [Formulas](#formulas)
6. [Updating](#updating)
7. [Troubleshooting](#troubleshooting)
8. [Project layout](#project-layout)
9. [Adding a new tool](#adding-a-new-tool)

## Installing

The easiest way is the one click installer. See the [install guide](https://chadisfred.github.io/Slidetools/manual/install.html), which is written to be handed to anyone who wants the add-in. The options below are for doing it by hand or for whole teams.

The install file is `slidetools-manifest.xml`. Save it somewhere permanent on your PC, for example `C:\Users\<you>\Slidetools\slidetools-manifest.xml`. Don't move or rename it afterwards, because PowerPoint looks for it at that exact path.

### Option 1: per user setting (no admin rights needed)

1. Close PowerPoint completely.
2. Open Command Prompt (not as administrator) and run, with your own path:

   ```
   reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v SlideTools /t REG_SZ /d "C:\Users\<you>\Slidetools\slidetools-manifest.xml" /f
   ```

3. Open PowerPoint. A Grid Panel button appears on the Home tab in a group called Slide Panel. If it doesn't, look under Home, Add-ins, More Add-ins, My Add-ins.

To remove it:

```
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v SlideTools /f
```

### Option 2: shared folder catalog (needs folder sharing rights)

1. Share the folder that holds the manifest and copy its network path, such as `\\YOURPC\Slidetools`.
2. In PowerPoint, go to File, Options, Trust Center, Trust Center Settings, Trusted Add-in Catalogs.
3. Paste the path into Catalog Url, click Add catalog, tick Show in Menu, click OK and restart PowerPoint.
4. On the Home tab, click Add-ins, then Advanced or More Add-ins, choose Shared Folder, pick Slide Tools Panel and click Add.

### For a whole team

Ask IT to deploy `slidetools-manifest.xml` from the Microsoft 365 admin center under Integrated apps. Everyone then gets the button without any local setup.

## Using the panel

A grid of tool buttons sits at the top of the panel, grouped by job:

| Row | Tools |
| --- | --- |
| Placing shapes | Grid, Position, Size, Copies, Animation |
| Text and colour | Typography, Colours, Text styles, Shape styles |
| Content | Images, Tables, Corners, Similar |

Click a button to open that tool. The other tools close, so there is little scrolling. Click the open tool's button again to close it.

1. Pin a tool with the pin in its header to keep it open while you open another, for example Size and Position together. Pinned tools show first and have a dot on their button.
2. The panel follows the selection. Selecting a picture opens Images, a table opens Tables, and clicking into text opens Typography. Selecting shapes never changes what is open. The matching button gets a dashed outline either way, and Open tools with the selection in the settings menu turns the opening off.
3. Show all tools in one list, in the settings menu, brings back the stacked view, where each tool folds and unfolds on its own.

Every tool header has the same controls:

| Control | What it does |
| --- | --- |
| Dotted grip on the left | Drag to change the tool's place in the button grid, or in the list |
| Tool name or the cross on the right | Close the tool. In the one list view it is an arrow that folds the tool |
| Pin | Keep the tool open when you open another |
| Detail button (Basic, More, Full) | Click to cycle how much of the tool is shown |
| i icon | Hover or tab to it for a short guide to the tool |
| Three dots | Move earlier, move later, reset the tool to its defaults, or hide it |

The sliders button at the top right shows the versions and updates, and lists every tool so you can show or hide each one. It also has Show all tools in one list, Open tools with the selection, Unpin all (or Unfold all and Fold all in the one list view), Share settings, Keyboard shortcuts, Ask before replacing a slide again and Reset panel arrangement.

The circular arrow next to it repeats the last action on the current selection, with the same values. Hover it to see what it will do. It covers Resize, Match size, Make copies, Align, Distribute, Even gaps, Snap, Round corners, and applying a style or colour.

Share settings copies a settings code for your text styles, shape styles, table styles, colours and grid presets. Send it to someone, and they paste it under Import, then choose Merge (add new items and update ones with the same name) or Replace. The code is plain text you copy and paste yourself.

Cropping pictures and making copies replace the slide with an edited copy. The first time, a message explains that slide comments may not carry over and Undo takes more than one step, with a Don't ask again box.

Detail levels:

1. Basic shows just enough for the everyday job.
2. More adds the common settings.
3. Full shows everything, including help text and less used options.

Tools that don't need three levels only offer the ones they use. Corners has Basic and Full. On a narrow panel the level button shows just its dots, and its tooltip names the level.

The panel remembers your values, tool order, the open and pinned tools, detail levels, folded panels and hidden tools between sessions. If PowerPoint uses a dark theme, the panel follows it.

## The tools

### Grid

Works out an evenly spaced grid and draws it on the slide, fits selected shapes into it, and snaps shapes onto it or onto your template's guides.

| Level | Shows |
| --- | --- |
| Basic | Presets, columns and rows, Draw on slide, Fit selection, Snap with a choice of Guides, Columns or Steps |
| More | Preview, cell size, grid spacing, area width and height, Use selection as area, gap sizes, ratios, snap options, template guide report |
| Full | Gap and grid tiles, left and top position, draw with rounded corners, group after drawing, save your own presets |

Tips:

1. Use selection as area: select a box or placeholder and the grid takes its position and size. With several shapes selected, it uses the box around all of them.
2. The preview shows selected shapes as dashed outlines, so you can see where they sit against the grid before you click Fit.
3. Fit places shapes in reading order: top to bottom, then left to right.
4. Your own presets appear after the built in ones. Hover a preset and click the small cross to delete it. Resetting the tool doesn't delete them.
5. After drawing, the new shapes are selected. With Group shapes after drawing on, they're grouped first. Both need a recent version of PowerPoint.
6. Choose None for a gap to put cells edge to edge.
7. Column ratio and Row ratio make unequal columns and rows. Type `2:1:1` for a first column twice as wide as the other two. The number of parts sets Columns or Rows, and changing the count by hand clears the ratio.
8. Snap moves each shape's top left corner to the nearest line, and a line under the buttons says what the chosen mode will do:
   1. Guides: the guides you see on the slide. That covers presentation guides (the ones you can drag on any slide, which a template usually brings with it) and guides in the current slide's layout and master. "Snap all four edges" also moves the right and bottom edges to the nearest guide.
   2. Columns: the column and row corners of the grid. "Resize to the cells it covers" also sizes the shape to span whole cells.
   3. Steps: every `g` from the grid area's top left. Can snap all four edges.
9. With Guides chosen, the More level reports the guides: how many there are, their spacing, any that sit slightly off the pattern, for example "8.34722 in (1 pt right)", and where they were found. Use for grid sets the grid spacing and the grid area's left and top from the guides at full precision. Read again refreshes them after you move, add or remove guides.

Guides are read from a read only copy of the open file, so they're never changed. The copy is taken once and kept until you click Read again. If PowerPoint won't hand over the file, the panel falls back to a copy of the slide (PowerPoint API 1.8). That fallback only has the layout and master guides, and the report says so.

If shapes drift a little further from the guides the further they are across the slide, the step size is slightly off. PowerPoint shows guide positions to two decimal places, so a 30 pt guide grid (0.41667 in) looks like 0.42. Use for grid, or typing `30pt` in Grid spacing, fixes it.

### Position

Lines shapes up, spaces them exactly, nudges them by the grid step, and copies one shape's position and size onto others.

| Level | Shows |
| --- | --- |
| Basic | Align buttons (left, centre, right, top, middle, bottom), distribute across or down with a gap, Nudge by an amount, Copy and Paste with Position, Size or Both. Nudge has [keyboard shortcuts](#keyboard-shortcuts) |
| More | Align to the selection, the grid area or the largest shape, gap result, even gaps, Paste from (anchor point) |
| Full | Swap two shapes |

Tips:

1. The gap and nudge boxes take formulas, such as `g`, `0.5g` or `gap`, using the grid's values.
2. Distribute keeps the first shape where it is and places the rest after it with the exact gap. Even gaps keeps the outer shapes still and shares the space between them equally.
3. Aligning a single shape to the selection aligns it to the grid area instead.
4. Copy position and size: select a shape and click Copy. The panel shows its position and size until it is closed. Go to another slide, select one or more shapes, choose Position, Size or Both, and click Paste.
5. At More, Paste from chooses which point lines up when the sizes differ. Centre places shapes on the copied shape's centre, and with Size only, keeps each shape's own centre still while it is resized.

### Size

Sets the width and height of every selected shape, using plain numbers or formulas, and matches sizes.

| Level | Shows |
| --- | --- |
| Basic | Width, Height, Resize |
| More | Live result, lock aspect ratio, reference point, Use selection, match to largest or smallest |
| Full | Recent formulas |

Hover over the small i in the tool's header, or tab to it, to read how formulas work.

Tips:

1. Press Enter in Width or Height to apply just that one.
2. With Lock aspect ratio on, the other side follows. If both boxes are filled, Resize uses Width.
3. The reference point sets which point stays still: a corner, an edge or the centre. Copies and the size shortcuts use the same point.
4. Match makes every selected shape the size of the largest or smallest one by area. Lines are matched to each other by length.
5. Recent formulas keeps your last six. Click one to put it back in the boxes.
6. At More, the line under the boxes shows the selected shape's size in inches and in steps of g, such as `6g x 2.5g`. Sizes that land on whole, half, quarter or eighth steps show exactly, such as `3.375g`, and others as about, for example `about 3.33g`.
7. Lines are resized by length, whichever box you type in. Horizontal, vertical and diagonal lines keep their direction.
8. With Crop pictures to fit on (in Images), a picture whose frame changes shape is cropped to fill it instead of being stretched. PowerPoint add-ins can't crop directly, so the slide is swapped for an identical copy with the new crop. This happens in memory inside the panel and nothing is sent anywhere. Slide comments may not carry over, and Undo takes more than one step. Turn it off to keep pictures in proportion instead.

### Copies

Makes spaced copies of the selected shapes in one step, like Illustrator's Transform effect.

| Level | Shows |
| --- | --- |
| Basic | Copies, Move across and down, Make copies |
| More | Rows, Space by, scale, rotate, reference point, copies preview and summary |
| Full | Reflect X and Y |

Tips:

1. Set how many, then Move across and Move down (formulas such as `w+gap`), and optionally Scale, Rotate and Reflect. Each copy is transformed from the one before, so a rotate of 45 with 7 copies makes a ring. Space by Gap measures the move as space between edges. Leave a Move box empty for no movement that way. If a box can't be read, such as `h+` with nothing after the plus, it turns red and the message says what to change.
2. Copies use the same in-memory slide swap as cropping, so they are exact copies of text, formatting, pictures, groups and tables. A copied placeholder becomes an ordinary shape, and tables and charts don't rotate or flip.
3. Rows turns copies into a grid. With 3 copies and 3 rows you get 4 across by 3 down. Move down is the step between rows, and left empty it is the shape's height plus the row gap.
4. The reference point is shared with Size, so changing it in either tool changes both.

### Images

Puts pictures into shapes like a picture placeholder, and moves or resets their crop.

| Level | Shows |
| --- | --- |
| Basic | Paste into shape, Crop position, Reset crop |
| More | Crop pictures to fit (used by Size) |

Tips:

1. Select a shape, such as a rounded card, click the paste box and press Ctrl+V. You can also drop a picture on the box or choose a file. The shape becomes a real picture with the same position, size, corners and outline, cropped to fill it without stretching.
2. The whole picture is kept, so PowerPoint's Crop button can pan or zoom it afterwards. Pasting onto a picture swaps its image and keeps its frame.
3. Crop position sets which part of the picture stays in view. With pictures selected, clicking a point moves their crop there. Reset crop shows the whole picture again.
4. The slide is swapped for an identical copy, in memory, to place or crop pictures, so slide comments may not carry over and Undo takes more than one step. Text in a shape is removed when a picture is placed in it.
5. Pictures are placed at full resolution and in their own format. A very large photo can make PowerPoint's Crop slow. Select it and use Picture Format, Compress Pictures to speed it up. Swapping a picture removes the old image from the slide, so it doesn't add to the file size.

### Animation

Smooth custom animations that PowerPoint's menus can't make: real easing curves, several effects at once, exact numbers, a stagger across shapes, and presets.

| Level | Shows |
| --- | --- |
| Basic | Presets (starters and yours), Entrance, Exit or Emphasis, preview with Play, Fade, Move (travel and distance), Scale, Rotate, or Pulse, Spin and Dim for emphasis, ease and direction with a curve, ease as text, duration, delay, Start, Apply to selection |
| More | Stagger and order, Replace their animations, preset name with Save and Capture, Remove animations from selection |

Tips:

1. Eases use GSAP's names and settings: `none`, `power1` to `power4`, `sine`, `expo`, `circ`, `back(overshoot)`, `elastic(amplitude, period)` and `bounce`, each with `.in`, `.out` or `.inOut`. Type one such as `back.out(1.7)` or pick it from the list. The curve is written to the slide as keyframes, so it plays the same as the preview.
2. Starters: Fade up, Pop, Slide in, Drop in, Zoom in, Swing in, Elastic grow, Rise out, Shrink out, Pulse and Gentle spin. Click one to fill the controls, then adjust.
3. With several shapes selected, the first one uses Start and the rest follow by the stagger, in the order you choose.
4. Save keeps the settings as a preset in the presentation. Capture keeps the selected shape's own animations, exactly as they are, so you can reuse an animation from any deck you can use.
5. Replace their animations, on by default, removes a shape's earlier animations first, so applying again doesn't stack effects.
6. PowerPoint add-ins can't add animations directly, so the slide is swapped for an identical copy with the animation, in memory. The result is a normal PowerPoint animation that shows in the Animation Pane and can be retimed there.

### Typography

Font, size, style and case, plus exact text spacing, margins and autofit, for the highlighted text or the selected shapes.

| Level | Shows |
| --- | --- |
| Basic | Font, Size with minus and plus, Bold, Italic, Underline, text colour, live preview, character spacing, line spacing in points or lines, space before and after, Apply spacing, margins, autofit, wrap |
| More | Case (All caps, Small caps, Strike), kerning, baseline, Read from selection, Live preview switch |

Tips:

1. Font, size, style and case apply straight away, to the highlighted text when there is some, otherwise to all the text in the selected shapes. The buttons light up for styles that are on, and half light when the text is mixed.
2. The font list starts with the theme's heading and body fonts, then fonts used on the slide. Type a name and press Enter, or pick one.
3. Every points box takes formulas with the grid names in points: `2g` in Size makes text two grid steps tall, `g/2` in a margin gives half a step, and `g` in Line spacing puts lines on the grid. Size also takes `+2` or `*1.5`, worked out from each text's own size. A box that can't be read turns red and says why.
4. The colour button shows the text colour and opens Colours with the target set to Text, keeping Typography pinned beside it.
5. Fill in only the spacing you want to change. Empty boxes are left as each shape has them. Each spacing box shows the value in use now, read from the shape or its template, and minus and plus step from it so the preview shows the result before you click Apply spacing.
6. Line spacing in points keeps the same distance whatever the font size, which helps text in different boxes line up.
7. Margins, autofit and wrap apply straight away. Spacing, kerning and baseline swap the slide for an identical copy, in memory, because add-ins can't set them directly, so Undo takes more than one step.
8. Read from selection copies a shape's settings into the boxes, ready to apply to others.
9. The live preview shows a scaled copy of the first selected text box with the values in the boxes, so you can see line breaks before applying. The dashed line marks the shape's height. A shape with no text shows no preview. Text that shrinks to fit is shrunk the same way in the preview. Fonts the panel can't use, such as Office cloud fonts, show as a similar font with a note.

### Corners

Sets the corner radius of rounded rectangles in points, so corners match across shapes of different sizes.

| Level | Shows |
| --- | --- |
| Basic | Radius, Round selection |
| Full | Radius presets, Read radius from selection, notes |

Press Enter in the Radius box to round the selection straight away. Square rectangles can't be converted from the panel. Draw them rounded, or use Round in the VBA Slide Tools. Corner changes need PowerPoint API 1.10 or later.

### Select similar

Finds shapes like the one you've selected, on this slide or slide by slide. Its button in the grid is called Similar.

| Level | Shows |
| --- | --- |
| Basic | Match by chips, Select similar on slide |
| More | Previous slide and Next slide with a count of matches |

1. Select an example shape, tick what to match by (Fill colour, Outline, Shape style, Text style, Size, Type), then click Select similar on slide. Every matching shape on that slide is selected.
2. PowerPoint can only select shapes on one slide at a time. Next slide and Previous slide jump to the next slide that has matches and select them there, wrapping round at the end. The status line shows how many matches are on that slide and in the whole presentation.
3. The example is remembered after you click Select similar, so you can keep stepping through slides after the selection changes. Changing the ticks applies to the next step.
4. Size allows half a point of difference. Pictures, tables and groups can only be matched by Type and Size.
5. Masters and layouts are never searched.

### Styles

Character styles in the spirit of InDesign: save the look of some text once, apply it anywhere, then change it everywhere at once. Styles are saved inside the presentation, so they travel with the file.

| Level | Shows |
| --- | --- |
| Basic | Your styles (click one to apply it), Sample, Update all |
| More | What style the selected text uses, each style's settings, New blank style |

A style can set font, size, colour, bold, italic, underline, strikethrough, caps (normal, all caps, small caps) and position (normal, superscript, subscript). Any setting can be left out with Leave. Every style needs a colour and at least a font or a size.

Fonts and weights:

1. The Font box suggests fonts used in this presentation first, then fonts in your styles, then your installed fonts. Typing any other name works too.
2. Listing installed fonts depends on the browser inside PowerPoint. The first time, it may ask for permission. If it's refused or unavailable, the panel remembers that for about 12 hours and doesn't ask again. Fonts from the presentation still appear.
3. Fonts that already carry a weight in their name, such as Segoe UI Semibold, Arial Black or Roboto Light, automatically get Bold set to Off, because PowerPoint's bold on top would thicken them artificially. You can switch it back on, with a warning.
4. When a style with a weighted font has Bold on Leave, applying it switches bold off on the text, so leftover bold doesn't stack.
5. A style with Bold On but no font gets a warning, because on a light or regular font that gives a fake bold. Pick the font's real bold version instead.

Making and using styles:

1. Select some text and click Sample. The panel fills in the settings from that text. Settings that are mixed in the selection are left out. Name it and click Save style. The sampled text is linked to the new style straight away.
2. To apply a style, select text, or whole text boxes, and click the style.
3. To change a style, click its pencil, edit and save. It shows Edited until you push the change.
4. Click Update all to push edits through the presentation. Nothing changes until you confirm.

Update all checks every normal slide. It never touches slide masters or layouts. For each style it reports:

1. Linked text, which will update. Settings you changed by hand on linked text, such as a different colour or italic, are kept.
2. Lookalikes: text that matches a style's font settings but isn't linked, usually because it was recoloured from the ribbon or written before the style existed. Choose to leave them alone, apply the style and its colour, or apply the style but keep their own colour. Review each one lists them by slide. Click an item to select it on the slide, or give it its own choice.

At the More level, selecting text shows its style. A "+" means it has local changes. You can reset it to the style, unlink it, or relink a lookalike while keeping its colour.

How linking works: add-ins can't attach data to individual characters, so each style's number is hidden in a tiny shift of the text colour, at most 3 out of 255 per channel, which can't be seen. Text boxes using a style also carry a hidden label. This means:

1. A style's exact hex code may differ slightly from the value you typed. #0071E3 might be stored on the text as #0070E1.
2. Recolouring styled text from the ribbon breaks its link. Update all then offers it as a lookalike, so nothing is silently lost.
3. Copying text with Format Painter or the eyedropper copies its style code too.
4. A presentation can hold up to 63 styles.
5. Colours are fixed values, not theme colours.
6. Paragraph settings such as space before and after, line spacing and kerning aren't available to add-ins, so styles cover character settings only.
7. Tables aren't checked by Update all.
8. Update all can be undone with Ctrl+Z, though it may take several presses.

### Shape styles

Named looks for shapes: fill, outline, corners and text box settings, optionally with one of your text styles for the text inside. Saved in the presentation, like text styles.

| Level | Shows |
| --- | --- |
| Basic | Your shape styles (click one to apply it to the selected shapes), Sample, Update all |
| More | What style the selected shape uses, each style's settings, New blank style |

A shape style can set:

| Area | Settings |
| --- | --- |
| Fill | Solid colour or no fill, transparency |
| Outline | Line or no line, colour, weight, dash, line type (single, double, thick and thin, triple), transparency |
| Corners | Radius in points, on rounded rectangles |
| Text box | Left, right, top and bottom margins, vertical alignment, wrap text, autofit |
| Text | One of your text styles, applied to all the text inside |

Any setting can be left out with Leave.

Using it:

1. Select a shape and click Sample. Name the style and click Save style. The sampled shape is linked and gets any changes you made before saving.
2. Select shapes and click a style to apply it, including its text style.
3. Edit a style with its pencil, then click Update all. It only updates shapes. Text styles have their own Update all in the Styles tool.
4. Update all lists linked shapes, which update while keeping settings you changed by hand, and lookalikes: unlinked shapes with the same fill and outline. For lookalikes choose to leave them, apply the style, or apply the style but keep their own fill and outline colours, for the whole group or one by one.
5. At the More level, selecting a shape shows its style, a "+" for local changes, and buttons to reset, unlink or link a lookalike.

Shapes are linked by a hidden label on the shape, so recolouring a shape never breaks its link. If a style's text style changes, Update all applies the new text style to linked shapes' text.

Limits:

1. Gradient, picture and pattern fills can't be set or sampled. Sampling leaves the fill out.
2. Shadows, glow, soft edges, 3D and arrowheads aren't available to add-ins.
3. Lines, pictures, charts and tables aren't included.
4. Only normal slides are changed, never masters or layouts.

### Tables

Inserts tables that are already formatted, or formats the selected table, using named table styles saved in the presentation.

| Level | Shows |
| --- | --- |
| Basic | Column roles for the selected table, Stack repeats, table styles, Insert table |
| More | Style editor, New style, Update all, the style to insert with, which style the table uses and Unlink |

Column roles. Select a table and each column appears as a chip. Click a chip to cycle its role:

1. Category: stacked labels such as a year or a team, shown once per group. Left aligned.
2. Label: row text that doesn't repeat, such as an item name. Left aligned.
3. Data: numbers. Right aligned, with a thin line on the left of the column.

The first column starts as a Category. Text columns next to it that repeat or leave blanks are guessed as Category, mostly numeric columns as Data, and the rest as Label. Roles are saved with the table.

The default table style:

1. A 1 pt black rule under the header row and under the table.
2. A 1 pt rule above each new group. A new value in a category column draws the rule from that column to the right edge, so a new year runs the full width and a new sub-group starts at its own column. With only one category column, these group lines are 1/4 pt instead, so only the header, total and bottom rules are heavy. The style editor can switch this off.
3. Total rows, found when the row's first filled cell starts with "Total", get Demi text and a 1 pt rule above, from the Total cell to the right edge.
4. 1/4 pt black lines between rows in label and data columns. Category columns only get a line where a group starts.
5. #F2F2F2 fill on the 2nd, 4th, 6th... columns, body rows only.
6. Fonts: header Avenir Next LT Pro with Bold on, first column Avenir Next LT Pro Demi, body Avenir Next LT Pro, 12 pt black.
7. Padding of 0.05 in (3.6 pt) on every side, except none on the first column's left and the last column's right, so text lines up with the table edges.
8. The header row sits at the bottom of its cells, body rows in the middle, and merged cells at the top.

Every part can be changed in the style editor: each part's font (or a text style from the Styles tool, which wins over the font settings), each line's weight, colour and visibility, the column fill, padding, and the word that marks total rows.

Tips:

1. Insert table adds a table at the grid area set in Grid, at its full width, with header text and roles set (Category, Sub-category, Item, then Value columns).
2. Clicking a style on a table that's already formatted reapplies it, keeping the table's roles. Changing a role reformats a linked table straight away.
3. Stack repeats blanks a category label when it repeats the row above within the same group. It never touches label or data columns or total rows, and PowerPoint's Undo brings the labels back.
4. Spaces, tabs and line breaks at the start and end of cells, common in text pasted from Excel, are trimmed before formatting so wrapping and column widths use the clean text. Only cells that change are rewritten. Turn it off in the style editor, or click Trim spaces for a one-off clean up.
4. Edit a style, then Update all reformats every table on normal slides that uses it.
5. Inserting needs PowerPoint API 1.8. Formatting an existing table needs 1.9. Only the top left cell of a merged area can be formatted.
6. Tables inside content placeholders (made with the table icon in a layout) work the same as other tables.
7. At the More level, a line under the column chips shows what PowerPoint reports about the selection, such as "1 selected shape (Placeholder holding Table), slide 3, 1 table on the slide". It helps when a table isn't being picked up.
8. Paste data: copy cells in Excel and paste them into the box under Insert table. Insert table then builds the table from them, with the first row as the header and column roles worked out from the values. Up to 40 rows and 20 columns.
9. Line up decimals, a style setting that is on by default, pads numbers in data columns to the same number of decimal places (1.5 next to 2.25 becomes 1.50), since add-ins can't set decimal tab stops.

Grid-based sizes. With the grid spacing as g:

1. The table fills the grid area width, rounded down to whole g.
2. Category and data columns are as wide as their longest text needs, rounded up to whole g. Label columns share the rest, also in whole g. With no label column, the spare room goes to the data columns.
3. Every row, header included, is 1g tall. A row whose text wraps grows in steps of 1/4 g.
4. The style editor can switch this off, turn off fitting to text (equal columns instead), or change the minimum row height and the step.

Check lines. Some PowerPoint versions show table lines the opposite way round, so lines that should show are hidden and the other way round. The first time a table is formatted, the panel adds a small test table with cells A and B and asks which one has a thick line under it. The answer is remembered on this computer and the test table is deleted. Run it again from Check lines at the More level after an Office update.

### Colours

Your template's colours and your own, one click away.

| Level | Shows |
| --- | --- |
| Basic | Fill, Outline or Text switch, and three tabs: Theme, Template and Yours |
| More | On the Yours tab, rename, reorder and remove your colours |

The tabs:

1. Theme: the ten theme colours in PowerPoint's order (Background 1, Text 1, Background 2, Text 2, Accent 1 to 6), with the same five rows of tints and shades as PowerPoint's colour menu. Hover a swatch for its name, such as "Accent 1, Lighter 40%".
2. Template: the Custom Colors saved in the template's theme. Single colours sit in one row, and each colour family (a colour with its tints) is a column, the same way the theme tints work. Empty white slots from PowerPoint's 10 column layout are left out. The grids keep their columns at any panel width, so nothing wraps.
3. Yours: your own colours, saved in the presentation. Add one by hex code, the colour picker, or Add from selection, which takes the selected shape's fill or the selected text's colour.

The panel remembers the last tab you used. Tabs for colours the deck doesn't have are hidden.

Notes:

1. Theme and template colours are read only. They belong to the slide master, so the panel never changes them. Template colours are read from a copy of the first slide, which needs PowerPoint API 1.8.
2. Colouring styled text from here keeps it linked to its text style, recording the new colour as a local change. The ribbon's font colour button breaks the link. Shapes keep their shape style link either way.
3. The colour pickers in the Styles and Shape styles editors suggest these colours.

## Keyboard shortcuts

These work while the slide has focus. On a Mac use Cmd+Option+Shift instead of Ctrl+Alt+Shift. They need PowerPoint for Windows 2601 or later, PowerPoint for Mac 16.105.2 or later, or PowerPoint on the web. The list is also in the panel's settings menu.

| Action | Keys |
| --- | --- |
| Nudge left, up, down, right | Ctrl+Alt+Shift+Left, Up, Down, Right |
| Wider, narrower, taller, shorter | Ctrl+Alt+Shift+L, J, I, K |
| Repeat last action | Ctrl+Alt+Shift+D |

Nudge and the size keys step by the Nudge by amount in Position, and sizes change around the reference point in Size. I, J, K and L are laid out like the arrow keys. Repeat covers panel actions only. To repeat a copy dragged with the mouse, use PowerPoint's Ctrl+D: duplicate, drag the copy into place, and each further Ctrl+D uses the same offset. Shortcuts were added in 2.25.0 with a change to the install file, so an install from before 2.25.0 needs the installer run once more.

The keys are declared in `shortcuts.json`, which the manifest links with `ExtendedOverrides`, and connected to the tools in `js/shortcuts.js`. Shortcuts need the shared runtime set in the manifest. Keep the two files in step when adding one. When the keys change, raise the `?v=` on the `shortcuts.json` address in the manifest and the manifest `<Version>`, so PowerPoint fetches the new list, and ask users to run the installer again.

## Formulas

Every number box accepts maths, for example `12.5-2*0.42`. Values are in inches unless you add a unit: `30pt`, `1.5cm` or `12mm`. Boxes keep five decimal places, so `30pt` becomes 0.41667. The Size, Copies and Position boxes also understand these names:

| Name | Meaning |
| --- | --- |
| `g` | Grid spacing |
| `col` | Width of the first column |
| `row` | Row height |
| `gap` | Horizontal gap |
| `vgap` | Vertical gap |
| `w` | The shape's current width |
| `h` | The shape's current height |

Examples: `1.5g`, `10g`, `2col+gap`, `*2` (double the current size), `1.4*2`. All sizes are in inches.

Typography's boxes are in points instead, and take the same names in points: with a 30 pt grid, `g` is 30 and `g/2` is 15. Its Size box also takes `+2` and `*1.5`.

## Updating

Changes to the panel itself go live without reinstalling:

1. Changes are made on a branch and opened as a pull request.
2. Merging the pull request into `main` publishes it through GitHub Pages within a minute or two.
3. Close and reopen the panel in PowerPoint.

Only changes to the ribbon button (its name, group or icons), the runtime or `shortcuts.json` links need an updated `slidetools-manifest.xml`. Replace your local copy, raise the `<Version>` number and restart PowerPoint.

Script and style links in `taskpane.html` end in `?v=` plus the version number. Raise that number with each release so PowerPoint doesn't keep old copies.

The panel reads its own version from the `?v=` on `js/update.js`, and its install version from `?install=` on the panel address in the manifest. `version.json` says what's current:

1. Each release sets `"panel"` to the new version, matching the `?v=` numbers.
2. A release that changes the manifest raises `"install"`, the `?install=` number on both panel addresses in the manifest and the manifest `<Version>`, and can say why in `"installNote"`.

The panel checks `version.json` when it opens. When it's out of date, a dot appears on the settings button, and the top of the settings menu offers Update panel (a reload) or Download installer, with What's new taken from this changelog.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| The panel shows an old version | Close PowerPoint, delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, reopen |
| The panel is blank or shows an error | Check the web address above loads in a browser, then reopen the panel |
| "Error loading add-ins" at launch | It can happen when the network is slow as PowerPoint starts. Restart PowerPoint. If it keeps happening, check the web address above loads in a browser |
| The Grid Panel button is missing | Run `reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer"` and check the path is right, then fully restart PowerPoint (end any PowerPoint entries in Task Manager) |
| A slide button says the version is too old | Update Office. Draw, Fit, Snap, Resize, Select similar and Styles need PowerPoint API 1.5. Grouping, caps, strikethrough and superscript in styles, and Update all inside groups need 1.8. Inserting tables needs 1.8. Formatting existing tables needs 1.9. Corners need 1.10 |
| Update all reports something you didn't expect | Click Cancel. Nothing changes until you click Update |

## Project layout

```
taskpane.html            Page shell: top bar, tool buttons, tool stack, script order
slidetools-manifest.xml  Office install file (ribbon button, panel address, shared runtime)
shortcuts.json           Keyboard shortcut keys, linked from the manifest
version.json             The current panel and install versions, read by the panel to offer updates
css/panel.css            All styling, light and dark themes, detail level rules
js/core.js               Shared state, saved settings, formula evaluator, grid maths, PowerPoint helpers
js/panels.js             Panel framework: tool buttons, pins, following the selection, headers, detail levels, folding, drag to reorder, menus
js/tools/snap.js         Snap, shown inside Grid (guides, columns, steps)
js/tools/layout.js       Grid tool (id layout)
js/tools/resize.js       Size tool (id resize: resize, lines, crop, match size, shared reference point)
js/tools/copies.js       Copies tool
js/animcore.js           Eases (GSAP-style names) and slide timing XML for animations
js/tools/animate.js      Animation tool (custom effects, stagger, presets, capture)
js/tools/images.js       Images tool (paste into shape, crop position, reset crop)
js/tools/typography.js   Typography tool (spacing, kerning, baseline, margins, autofit)
js/rebuild.js            In-memory slide rebuild used to crop pictures and make copies
js/crop.js               Picture cropping for Size
js/clone.js              Copy maths and slide edits for Copies
js/shortcuts.js          Keyboard shortcuts: names each action and connects it to the tools
js/update.js             Versions and updates in the settings menu
js/tools/corners.js      Corners tool
js/tools/arrange.js      Position tool (id arrange: align, distribute, nudge, swap, copy and paste position)
js/tools/match.js        Select similar tool (id match)
js/tools/styles.js       Styles tool (character styles)
js/tools/shapestyles.js  Shape styles tool
js/tools/tables.js       Tables tool (table styles, insert and format)
js/tools/colours.js      Colours tool (palette)
js/app.js                Starts everything and connects to PowerPoint
assets/                  Ribbon icons
install/                 One click installer and uninstaller for Windows
INSTALL.md               Short pointer to the install guide in the manual
```

There is no build step. GitHub Pages serves the files as they are.

## Adding a new tool

1. Create `js/tools/yourtool.js` and register the tool:

   ```js
   (function () {
     var ST = window.ST;
     var defaults = { myValue: "1" };
     ST.addDefaults(defaults);

     ST.registerTool({
       id: "yourtool",            // unique, used to save its order and level
       title: "Your tool",
       icon: '<svg viewBox="0 0 18 18">...</svg>',
       levels: [1, 2, 3],         // any of 1 Basic, 2 More, 3 Full
       defaultLevel: 1,
       defaults: defaults,        // what "Reset to defaults" restores
       tip: { what: "One line saying what the tool does.", steps: ["First step.", "Second step."] },
       build: function (body) {
         body.appendChild(ST.html(
           '<label class="field"><span class="label-row">My value' + ST.info("One short sentence about this setting.") + '</span>' +
           '<input type="text" id="myValue"></label>'));
         ST.bindField(ST.$("myValue"), "myValue", { number: true });
       }
     });
   })();
   ```

2. Add a `<script>` line for it in `taskpane.html`, after the other tools and before `js/app.js`.

3. Control what each level shows with attributes on any element:

   | Attribute | Shown at |
   | --- | --- |
   | none | Every level |
   | `data-min="2"` | More and Full |
   | `data-min="3"` | Full only |
   | `data-max="1"` | Basic only |

4. Useful helpers on `window.ST`: `state`, `changed()`, `on("change" | "selection" | "reset" | "layout", fn)`, `selection`, `withSelection(props, task)`, `run(task)`, `toast(message, isBad)`, `supports(version)`, `evaluate(text, vars)`.

5. Add a manual page `manual/<id>.html` named after the tool's `id`, so How to use in the tool's menu opens it. Follow the other tool pages: overview, annotated panel with a reference row (`data-control`) for every control, how it works, recipes, limits, messages and related tools. Link it from `manual.html` and the sidebar on every page.

New tools appear at the end of the tool buttons for anyone who already has a saved arrangement. To give a tool a place in the first order, add its `id` to `DEFAULT_ORDER` in `js/panels.js`.

### Writing panel text

1. No paragraphs in the panel. Any explanation goes behind an i icon (`ST.info(text)`).
2. The header tip (`tip`) is one line saying what the tool does, then up to 3 short steps, about 40 words at most.
3. Inline tips cover one setting each, in one or two short sentences.
4. Visible text is only labels, buttons and live status, such as "3 matches on slide 4". Empty states are one short line saying what to do.
5. Buttons start with a verb. Hover titles are 3 to 8 words. Use the same terms everywhere: grid area, g, category, label, data.

See [CHANGELOG.md](CHANGELOG.md) for the history of changes.
