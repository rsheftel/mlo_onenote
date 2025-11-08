# Convert between OneNote and MLO using the OPML format.

## To run the script version:

python OneNote_to_MLO.py

## To run the GUI version:

python OneNote_to_MLO_gui.py

## To build a stand alone executable:

1. Create the folder and copy the three files

mkdir onenote-to-mlo

cd onenote-to-mlo

(paste the three files from above)

2. Resolve dependencies + build

uv sync          # installs everything listed in pyproject.toml

uv run ./build_exe.py

3. The output will be here ./dist/onenote-to-mlo.exe
