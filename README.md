# Convert between OneNote and MLO using the OPML format.

This simple python GUI, with included .exe executable, will convert between OneNote and MLO using the OPML format.
The options are:

1. Save OneNote as .MHT and then convert to .OPML for import to MLO
2. Copy the text of a bullet list from OneNote paste into the GUI and convert to .OPML for import to MLO
3. Export from MLO to .OPML and conver to HTML then import into OneNote

When a task in MLO has a note, it will be added to the OneNote bullet list preceeded with #note=

When converting from text for .MHT file, any text after #note= will be converted to a note in the .OPML file.

## To run the GUI version

- python gui.py

## Windows Shortcut
- add a shortcut to the Windows start or desktop by copying MLO-to-OneNote.lnk and modifying the target.

## To build a stand alone executable

- python build_exe.py

- The output will be in ./dist/onenote-to-mlo.exe
