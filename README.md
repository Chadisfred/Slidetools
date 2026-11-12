# Slide Tools

A PowerPoint side panel for laying out grids, resizing shapes with formulas, rounding corners and managing character styles. It runs as an Office web add-in and is hosted on GitHub Pages:

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

Hover over the small i at the top right of the tool, or tab to it, to read how formulas work.

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

## Formulas

Every number box accepts maths, for example `12.5-2*0.42`. The Resize boxes also understand these names:

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
| A slide button says the version is too old | Update Office. Draw, Fit, Resize and Styles need PowerPoint API 1.5. Grouping, caps, strikethrough and superscript in styles, and Update all inside groups need 1.8. Corners need 1.10 |
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
js/tools/styles.js       Styles tool (character styles)
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
