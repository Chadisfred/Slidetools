# Installing Slide Tools

Slide Tools is a side panel for PowerPoint with layout grids, formula based resizing, alignment and snapping, character and shape styles, colours and more.

What you need:

1. PowerPoint for Windows, from Microsoft 365 or Office 2016 or later.
2. An internet connection. The panel loads from the web each time PowerPoint opens it.
3. No admin rights. Everything installs for your Windows account only.

## Quick install (recommended)

1. Close PowerPoint.
2. Download the installer: [Install-SlideTools.cmd](https://chadisfred.github.io/Slidetools/install/Install-SlideTools.cmd)
3. Open your Downloads folder and double click `Install-SlideTools.cmd`.
   - If your browser says the file isn't commonly downloaded, choose Keep.
   - If Windows shows "Windows protected your PC", click More info, then Run anyway.
4. A window opens, installs Slide Tools and says when it's done. Press any key to close it.
5. Open PowerPoint. The Grid Panel button is on the Home tab, in a group called Slide Panel. Click it to open the panel.

If the button isn't on the Home tab, go to Home, Add-ins, then My Add-ins, and pick Slide Tools from there.

The installer does two things:

1. It saves the add-in's install file in a `Slidetools` folder in your user folder.
2. It adds one PowerPoint setting that points to that file.

Leave that folder where it is, because PowerPoint reads the file from there.

## Manual install

Use this if the installer is blocked on your PC.

1. Close PowerPoint.
2. Press Windows + R, type `%USERPROFILE%` and press Enter. Make a new folder there called `Slidetools`.
3. Open [slidetools-manifest.xml](https://chadisfred.github.io/Slidetools/slidetools-manifest.xml) in your browser. Press Ctrl + S and save it in the new `Slidetools` folder, keeping the name `slidetools-manifest.xml`.
4. Press Windows + R, type `cmd` and press Enter. Paste this line into the black window and press Enter:

   ```
   reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v SlideTools /t REG_SZ /d "%USERPROFILE%\Slidetools\slidetools-manifest.xml" /f
   ```

   It should say "The operation completed successfully."
5. Open PowerPoint and look for Grid Panel on the Home tab.

## Updates

There's nothing to do. Slide Tools updates itself, and restarting PowerPoint picks up the latest version.

## Uninstall

Close PowerPoint, then download and double click [Uninstall-SlideTools.cmd](https://chadisfred.github.io/Slidetools/install/Uninstall-SlideTools.cmd).

Or remove it by hand. Run this line in a command window, then delete the `Slidetools` folder in your user folder:

```
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v SlideTools /f
```

## Troubleshooting

| Problem | What to try |
| --- | --- |
| No Grid Panel button | Check Home, Add-ins, My Add-ins. If it isn't there either, close PowerPoint completely, including any PowerPoint entries in Task Manager, and open it again. |
| The panel shows an old version | Close PowerPoint. Press Windows + R, paste `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, press Enter, delete everything in that folder, and reopen PowerPoint. |
| The panel is blank or shows an error | Check that [the panel's web address](https://chadisfred.github.io/Slidetools/taskpane.html) opens in your browser, then reopen the panel. |
| The installer can't download the file | A network filter may block it. Use the manual install steps. |
| A message says add-ins are blocked by your organisation | Your IT team controls add-ins. Ask them to deploy `slidetools-manifest.xml` from the Microsoft 365 admin center, under Integrated apps. |
| A slide button says PowerPoint is too old | Update Office from File, Account, Update Options. |

For a full guide to the tools, see the [README](README.md).
