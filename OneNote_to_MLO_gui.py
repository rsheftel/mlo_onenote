#!/usr/bin/env python
"""
GUI OneNote .mht → MyLifeOrganized OPML converter
(no drag-and-drop – uses normal file dialogs)
"""

from __future__ import annotations

import re
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, List

import quopri
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup, NavigableString, Tag
from loguru import logger


# --------------------------------------------------------------------------- #
# 1. MIME / quoted-printable extraction
# --------------------------------------------------------------------------- #
def extract_html_from_mht(mht_path: Path) -> str:
    raw = mht_path.read_text(encoding="utf-8", errors="replace")
    boundary_match = re.search(r'boundary\s*=\s*["\']?([^"\']*)["\']?', raw, re.I | re.M)
    if boundary_match:
        boundary = boundary_match.group(1)
        parts = re.split(rf"--{re.escape(boundary)}", raw)
        for part in parts[1:-1]:
            part = part.strip()
            if not part:
                continue
            header_body = re.split(r"\r?\n\r?\n|\n\n", part, maxsplit=1)
            if len(header_body) < 2:
                continue
            hdr, body = header_body
            if "text/html" in hdr.lower():
                if "quoted-printable" in hdr.lower():
                    body = quopri.decodestring(body.encode("utf-8")).decode("utf-8")
                return body
    # fallback – single part
    header_body = re.split(r"\r?\n\r?\n|\n\n", raw, maxsplit=1)
    if len(header_body) < 2:
        return raw
    hdr, body = header_body
    if "quoted-printable" in hdr.lower():
        body = quopri.decodestring(body.encode("utf-8")).decode("utf-8")
    return body


# --------------------------------------------------------------------------- #
# 2. Text cleaning – no &#10; and collapse whitespace
# --------------------------------------------------------------------------- #
def clean_text(text: str) -> str:
    text = re.sub(r"&#\d+;", " ", text)  # &#10; → space
    text = re.sub(r"\s+", " ", text)  # collapse all whitespace
    return text.strip()


# --------------------------------------------------------------------------- #
# 3. Parse OneNote lists (handles the weird <ol><ol><li> nesting)
# --------------------------------------------------------------------------- #
def parse_list(list_tag: Tag) -> List[dict[str, Any]]:
    if list_tag.name not in {"ol", "ul"}:
        raise ValueError(f"Expected <ol> or <ul>, got <{list_tag.name}>")

    tasks: List[dict[str, Any]] = []
    current_task: dict[str, Any] | None = None

    for child in list_tag.children:
        if isinstance(child, NavigableString):
            continue
        if child.name == "li":
            name = clean_text(child.get_text(separator=" ", strip=True))
            current_task = {"name": name, "subtasks": []}
            tasks.append(current_task)
        elif child.name in {"ol", "ul"} and current_task:
            current_task["subtasks"].extend(parse_list(child))

    return tasks


# --------------------------------------------------------------------------- #
# 4. Build OPML
# --------------------------------------------------------------------------- #
def build_outline(task: dict[str, Any]) -> ET.Element:
    el = ET.Element("outline")
    el.set("text", task["name"])
    for sub in task["subtasks"]:
        el.append(build_outline(sub))
    return el


def build_opml(task_list: List[dict[str, Any]]) -> ET.ElementTree:
    opml = ET.Element("opml", version="1.0")
    head = ET.SubElement(opml, "head")
    ET.SubElement(head, "title").text = "OneNote → MLO"
    body = ET.SubElement(opml, "body")
    for t in task_list:
        body.append(build_outline(t))
    return ET.ElementTree(opml)


# --------------------------------------------------------------------------- #
# 5. Core conversion
# --------------------------------------------------------------------------- #
def convert(mht_path: Path, opml_path: Path) -> None:
    html = extract_html_from_mht(mht_path)
    soup = BeautifulSoup(html, "lxml")

    body = soup.body
    if not body:
        raise ValueError("No <body> in the HTML part.")
    main_list = body.find("ol") or body.find("ul")
    if not main_list:
        raise ValueError("No <ol> or <ul> found – is this a numbered list page?")

    hierarchy = parse_list(main_list)
    if not hierarchy:
        raise ValueError("List found but no tasks extracted.")

    tree = build_opml(hierarchy)
    ET.indent(tree, space="  ", level=0)
    tree.write(
        opml_path,
        encoding="utf-8",
        xml_declaration=True,
        method="xml",
        short_empty_elements=False,
    )


# --------------------------------------------------------------------------- #
# 6. GUI (plain tkinter – no drag-and-drop)
# --------------------------------------------------------------------------- #
class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("OneNote → MLO OPML")
        self.geometry("560x220")
        self.resizable(False, False)

        pad = dict(padx=10, pady=8)

        # ---- Input -------------------------------------------------------
        ttk.Label(self, text="Input .mht file:").grid(row=0, column=0, sticky="w", **pad)
        self.ent_input = ttk.Entry(self, width=50)
        self.ent_input.grid(row=0, column=1, **pad)
        ttk.Button(self, text="Browse…", command=self.browse_input).grid(row=0, column=2, **pad)

        # ---- Output ------------------------------------------------------
        ttk.Label(self, text="Output .opml file:").grid(row=1, column=0, sticky="w", **pad)
        self.ent_output = ttk.Entry(self, width=50)
        self.ent_output.grid(row=1, column=1, **pad)
        ttk.Button(self, text="Save as…", command=self.browse_output).grid(row=1, column=2, **pad)

        # ---- Convert -----------------------------------------------------
        ttk.Button(self, text="Convert →", command=self.run_convert).grid(row=2, column=1, pady=20)

        # auto-fill output when input changes
        self.ent_input.bind("<KeyRelease>", self.auto_output)

    # ------------------------------------------------------------------- #
    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select OneNote .mht file",
            filetypes=[("MHTML files", "*.mht *.mhtml"), ("All files", "*.*")],
        )
        if path:
            self.ent_input.delete(0, tk.END)
            self.ent_input.insert(0, path)
            self.auto_output()

    def browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save OPML file",
            defaultextension=".opml",
            filetypes=[("OPML files", "*.opml"), ("XML files", "*.xml")],
        )
        if path:
            self.ent_output.delete(0, tk.END)
            self.ent_output.insert(0, path)

    def auto_output(self, *_) -> None:
        inp = Path(self.ent_input.get())
        if inp.suffix.lower() in {".mht", ".mhtml"} and inp.exists():
            out = inp.with_suffix(".opml")
            self.ent_output.delete(0, tk.END)
            self.ent_output.insert(0, str(out))

    def run_convert(self) -> None:
        in_path = Path(self.ent_input.get())
        out_path = Path(self.ent_output.get())

        if not in_path.exists():
            messagebox.showerror("Error", f"Input file not found:\n{in_path}")
            return
        if out_path.exists():
            if not messagebox.askyesno("Overwrite?", f"File exists:\n{out_path}\nOverwrite?"):
                return

        try:
            convert(in_path, out_path)
            messagebox.showinfo(
                "Success",
                f"OPML created!\n{out_path}\n\nImport into MyLifeOrganized → File → Import → OPML",
            )
        except Exception as exc:
            logger.exception("Conversion failed")
            messagebox.showerror("Conversion failed", str(exc))


# --------------------------------------------------------------------------- #
# 7. Entry point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    logger.remove()
    logger.add(sys.stderr, level="INFO")
    app = ConverterApp()
    app.mainloop()
