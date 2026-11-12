# Slide Tools

A PowerPoint side panel for laying out grids, resizing shapes with formulas and rounding corners. It runs as an Office web add-in and is hosted on GitHub Pages:

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
| More | Live result, lock aspect ratio, Use selection, match to largest or smallest |
| Full | Resize from point (anchor), recent formulas, formula help |

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
| A slide button says the version is too old | Update Office. Draw, Fit and Resize need PowerPoint API 1.5, grouping needs 1.8 and corners need 1.10 |

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
