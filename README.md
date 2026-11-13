# Slide Tools

A PowerPoint side panel for laying out grids, resizing shapes with formulas, rounding corners, selecting and matching similar shapes, and managing character and shape styles. It runs as an Office web add-in and is hosted on GitHub Pages:

https://chadisfred.github.io/Slidetools/taskpane.html

Open that address in a browser to try the panel in preview mode. The calculator works there, and the slide buttons work once the panel is opened from PowerPoint.

Note the capital S in `Slidetools`. The address doesn't work with a lowercase s.

## Contents

1. [Installing](#installing)
2. [Using the panel](#using-the-panel)
3. [The tools](#the-tools)
4. [Formulas](#formulas)
5. [Updating](#updating)
6. [Troubleshooting](#troubleshooting)
7. [Project layout](#project-layout)
8. [Adding a new tool](#adding-a-new-tool)

## Installing

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

Each tool sits in its own panel, stacked from top to bottom. Every panel header has the same controls:

| Control | What it does |
| --- | --- |
| Dotted grip on the left | Drag up or down to reorder the tools |
| Tool name or the arrow on the right | Fold the panel down to just its header, or unfold it |
| Detail button (Basic, More, Full) | Click to cycle how much of the tool is shown |
| i icon | Hover or tab to it for a short guide to the tool |
| Three dots | Move up, move down, reset the tool to its defaults, or hide it |

The sliders button at the top right lists every tool so you can show or hide each one. It also has Unfold all, Fold all and Reset panel arrangement.

Detail levels:

1. Basic shows just enough for the everyday job.
2. More adds the common settings.
3. Full shows everything, including help text and less used options.

Tools that don't need three levels only offer the ones they use. Corners has Basic and Full.

The panel remembers your values, tool order, detail levels, folded panels and hidden tools between sessions. If PowerPoint uses a dark theme, the panel follows it.

## The tools

### Layout

Works out an evenly spaced grid and draws it on the slide, or fits selected shapes into it.

| Level | Shows |
| --- | --- |
| Basic | Presets, columns and rows, Draw on slide, Fit selection |
| More | Preview, cell size, grid spacing, area width and height, Use selection as area, gap sizes |
| Full | Gap and grid tiles, left and top position, draw with rounded corners, group after drawing, save your own presets |

Tips:

1. Use selection as area: select a box or placeholder and the grid takes its position and size. With several shapes selected, it uses the box around all of them.
2. The preview shows selected shapes as yellow dashed outlines, so you can see where they sit against the grid before you click Fit.
3. Fit places shapes in reading order: top to bottom, then left to right.
4. Your own presets appear after the built in ones. Hover a preset and click the small cross to delete it. Resetting the tool doesn't delete them.
5. After drawing, the new shapes are selected. With Group shapes after drawing on, they're grouped first. Both need a recent version of PowerPoint.

### Resize

Sets the width and height of every selected shape, using plain numbers or formulas.

| Level | Shows |
| --- | --- |
| Basic | Width, Height, Resize |
| More | Live result, lock aspect ratio, resize from point (anchor), Use selection, match to largest or smallest |
| Full | Recent formulas |

Hover over the small i in the tool's header, or tab to it, to read how formulas work.

Tips:

1. Press Enter in Width or Height to apply just that one.
2. With Lock aspect ratio on, the other side follows. If both boxes are filled, Resize uses Width.
3. Resize from sets which point stays still: a corner, an edge or the centre.
4. Match makes every selected shape the size of the largest or smallest one by area.
5. Recent formulas keeps your last six. Click one to put it back in the boxes.

### Corners

Sets the corner radius of rounded rectangles in points, so corners match across shapes of different sizes.

| Level | Shows |
| --- | --- |
| Basic | Radius, Round selection |
| Full | Radius presets, Read radius from selection, notes |

Press Enter in the Radius box to round the selection straight away. Square rectangles can't be converted from the panel. Draw them rounded, or use Round in the VBA Slide Tools. Corner changes need PowerPoint API 1.10 or later.

### Arrange

Lines shapes up and spaces them exactly, using the same grid as Layout.

| Level | Shows |
| --- | --- |
| Basic | Align buttons (left, centre, right, top, middle, bottom), distribute across or down with a gap, Snap with a choice of Guides, Columns or Steps |
| More | Align to the selection, the grid area or the largest shape, gap result, even gaps, snap options, template guide report |
| Full | Nudge by an amount, swap two shapes |

Tips:

1. The gap and nudge boxes take formulas, such as `g`, `0.5g` or `gap`, using the Layout grid's values.
2. Distribute keeps the first shape where it is and places the rest after it with the exact gap. Even gaps keeps the outer shapes still and shares the space between them equally.
3. Aligning a single shape to the selection aligns it to the grid area instead.
4. Snap moves each shape's top left corner to the nearest line, and a line under the buttons says what the chosen mode will do:
   1. Guides: the guides built into your template, read from the current slide's layout and master. "Snap all four edges" also moves the right and bottom edges to the nearest guide.
   2. Columns: the column and row corners of the Layout grid. "Resize to the cells it covers" also sizes the shape to span whole cells.
   3. Steps: every `g` from the grid area's top left. Can snap all four edges.
5. With Guides chosen, the More level reports the template's guides: how many there are, their spacing, and any that sit slightly off the pattern, for example "8.34722 in (1 pt right)". Use in Layout sets the Layout spacing and the grid area's left and top from the guides at full precision. Read again refreshes them after you change the template.

Template guides are only read, never changed, and need PowerPoint API 1.8. Guides you drag onto ordinary slides yourself may not be included, because only the slide's layout and master are read.

If shapes drift a little further from the guides the further they are across the slide, the step size is slightly off. PowerPoint shows guide positions to two decimal places, so a 30 pt guide grid (0.41667 in) looks like 0.42. Use in Layout, or typing `30pt` in Spacing, fixes it.

### Match

Finds shapes like the one you've selected, and copies one shape's position and size onto others.

| Level | Shows |
| --- | --- |
| Basic | Match by chips, Select similar on slide, Copy and Paste with Position, Size or Both |
| More | Previous slide and Next slide with a count of matches, Paste from (anchor point) |

Select similar:

1. Select an example shape, tick what to match by (Fill colour, Outline, Shape style, Text style, Size, Type), then click Select similar on slide. Every matching shape on that slide is selected.
2. PowerPoint can only select shapes on one slide at a time. Next slide and Previous slide jump to the next slide that has matches and select them there, wrapping round at the end. The status line shows how many matches are on that slide and in the whole presentation.
3. The example is remembered after you click Select similar, so you can keep stepping through slides after the selection changes. Changing the ticks applies to the next step.
4. Size allows half a point of difference. Pictures, tables and groups can only be matched by Type and Size.
5. Masters and layouts are never searched.

Copy position and size:

1. Select a shape and click Copy. The panel shows its position and size until it is closed.
2. Go to another slide, select one or more shapes, and click Paste. Choose Position, Size or Both first.
3. At More, Paste from chooses which point lines up when the sizes differ. Centre places shapes on the copied shape's centre, and with Size only, keeps each shape's own centre still while it is resized.

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

## Formulas

Every number box accepts maths, for example `12.5-2*0.42`. Values are in inches unless you add a unit: `30pt`, `1.5cm` or `12mm`. Boxes keep five decimal places, so `30pt` becomes 0.41667. The Resize boxes also understand these names:

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

## Updating

Changes to the panel itself go live without reinstalling:

1. Changes are made on a branch and opened as a pull request.
2. Merging the pull request into `main` publishes it through GitHub Pages within a minute or two.
3. Close and reopen the panel in PowerPoint.

Only changes to the ribbon button (its name, group or icons) need an updated `slidetools-manifest.xml`. Replace your local copy, raise the `<Version>` number and restart PowerPoint.

Script and style links in `taskpane.html` end in `?v=` plus the version number. Raise that number with each release so PowerPoint doesn't keep old copies.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| The panel shows an old version | Close PowerPoint, delete the contents of `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, reopen |
| The panel is blank or shows an error | Check the web address above loads in a browser, then reopen the panel |
| The Grid Panel button is missing | Run `reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer"` and check the path is right, then fully restart PowerPoint (end any PowerPoint entries in Task Manager) |
| A slide button says the version is too old | Update Office. Draw, Fit, Resize, Match and Styles need PowerPoint API 1.5. Grouping, caps, strikethrough and superscript in styles, and Update all inside groups need 1.8. Corners need 1.10 |
| Update all reports something you didn't expect | Click Cancel. Nothing changes until you click Update |

## Project layout

```
taskpane.html            Page shell: top bar, tool stack, script order
slidetools-manifest.xml  Office install file (ribbon button and panel address)
css/panel.css            All styling, light and dark themes, detail level rules
js/core.js               Shared state, saved settings, formula evaluator, grid maths, PowerPoint helpers
js/panels.js             Panel framework: headers, detail levels, folding, drag to reorder, menus
js/tools/layout.js       Layout tool
js/tools/resize.js       Resize tool
js/tools/corners.js      Corners tool
js/tools/arrange.js      Arrange tool (align, distribute, snap)
js/tools/match.js        Match tool (select similar, copy position and size)
js/tools/styles.js       Styles tool (character styles)
js/tools/shapestyles.js  Shape styles tool
js/tools/colours.js      Colours tool (palette)
js/app.js                Starts everything and connects to PowerPoint
assets/                  Ribbon icons
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
       build: function (body) {
         body.appendChild(ST.html(
           '<label class="field">My value<input type="text" id="myValue"></label>' +
           '<div class="help" data-min="3">Only shown at Full.</div>'));
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

New tools appear at the bottom of the stack for anyone who already has a saved arrangement.

See [CHANGELOG.md](CHANGELOG.md) for the history of changes.
