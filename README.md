# Convert between OneNote and MLO using the OPML format.

## To run the script versions

python OneNote_to_MLO.py
python MLO_to_OneNote.py

## To run the GUI version

python OneNote_to_MLO_gui.py

add a shortcut to the Windows start or desktop by copying MLO-to-OneNote.lnk

## To build a stand alone executable

1. Create the folder and copy the three files

mkdir onenote-to-mlo

cd onenote-to-mlo

(paste the three files from above)

2. Resolve dependencies + build

uv sync          # installs everything listed in pyproject.toml

uv run ./build_exe.py

3. The output will be here ./dist/onenote-to-mlo.exe
