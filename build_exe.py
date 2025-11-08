#!/usr/bin/env python
"""
Build standalone .exe with uv + PyInstaller (robust version)
"""

import subprocess
import sys
from pathlib import Path

def main() -> None:
    script = Path(__file__).with_name("onenote_to_mlo_gui.py")
    if not script.exists():
        print(f"Error: {script} not found")
        sys.exit(1)

    cmd = [
        "uv", "run", "pyinstaller",
        "--onefile",
        "--noconsole",
        "--name", "OneNote-to-MLO",
        "--clean",                    # remove old build files
        str(script),
    ]

    # Optional icon
    icon = Path("icon.ico")
    if icon.exists():
        cmd.extend(["--icon", str(icon)])

    print("Building executable (this may take 30–60 seconds)…")
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("Success! → dist/OneNote-to-MLO.exe")
        print("\nTip: Test it on a clean machine (no Python installed)")
    except subprocess.CalledProcessError as e:
        print("PyInstaller failed:")
        print(e.stdout)
        print(e.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
