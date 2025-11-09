#!/usr/bin/env python
"""
GUI Converter: OneNote ↔ MyLifeOrganized

- OneNote (.mht) → MLO (.opml)
- MLO (.opml) → OneNote (.html)
"""

from __future__ import annotations

import quopri
import re
import sys
import tkinter as tk
import xml.etree.ElementTree as ET
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, List, Optional

from bs4 import BeautifulSoup, NavigableString, Tag
from loguru import logger


# --------------------------------------------------------------------------- #
# Common: Text cleaning
# --------------------------------------------------------------------------- #
def clean_text(text: str) -> str:
    text = re.sub(r"&#\d+;", " ", text)  # &#10; → space
    text = re.sub(r"\s+", " ", text)  # collapse all whitespace
    return text.strip()


# --------------------------------------------------------------------------- #
# Direction 1: OneNote (.mht) → MLO (.opml)
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


def convert_mht_to_opml(mht_path: Path, opml_path: Path) -> None:
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
# Direction 2: MLO (.opml) → OneNote (.html)
# --------------------------------------------------------------------------- #
def parse_opml_outline(outline_elem: ET.Element) -> Optional[dict]:
    """
    Recursively parse an <outline> element into a task dictionary.

    Skips the entire subtree if _status == "checked".
    """
    if outline_elem.get("_status") == "checked":
        return None

    name = outline_elem.get("text", "(no text)").strip()
    note = outline_elem.get("_note", "").strip()

    subtasks = []
    for sub in outline_elem:
        sub_task = parse_opml_outline(sub)
        if sub_task is not None:
            subtasks.append(sub_task)

    return {"name": name, "note": note, "subtasks": subtasks}


def parse_opml_body(body_elem: ET.Element) -> List[dict]:
    tasks = []
    for outline in body_elem.findall("outline"):
        task = parse_opml_outline(outline)
        if task is not None:
            tasks.append(task)
    return tasks


NUMBERING_STYLES = [
    "decimal",        # 1., 2., 3.
    "lower-alpha",    # a., b., c.
    "lower-roman",    # i., ii., iii.
    "decimal)",       # 1), 2), 3)
    "decimal",        # 1., 2., 3. (repeat)
]


def build_html_li(task: dict, soup: BeautifulSoup, level: int = 0) -> Tag:
    li = soup.new_tag("li")
    span = soup.new_tag("span", style="font-family:Calibri;font-size:11pt")
    span.string = task["name"]
    li.append(span)

    # Add note as italic sub-item
    if task["note"]:
        note_li = soup.new_tag("li")
        note_span = soup.new_tag("span", style="font-family:Calibri;font-size:11pt;font-style:italic")
        note_span.string = task["note"]
        note_li.append(note_span)
        ul = soup.new_tag("ul", style="list-style-type:none;margin-left:36pt")
        ul.append(note_li)
        li.append(ul)

    if task["subtasks"]:
        style_idx = level % len(NUMBERING_STYLES)
        list_style = NUMBERING_STYLES[style_idx]
        ol = soup.new_tag("ol", style=f"list-style-type:{list_style};margin-left:36pt")
        for sub in task["subtasks"]:
            ol.append(build_html_li(sub, soup, level + 1))
        li.append(ol)

    return li


def build_html(task_list: List[dict]) -> BeautifulSoup:
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>MLO → OneNote</title>
        <style>
            body { font-family: Calibri, sans-serif; font-size: 11pt; }
            ol { margin-left: 36pt; }
        </style>
    </head>
    <body>
    </body>
    </html>
    """
    soup = BeautifulSoup(html, "html.parser")

    ol = soup.new_tag("ol", style="list-style-type:decimal;margin-left:36pt")
    for task in task_list:
        ol.append(build_html_li(task, soup, level=0))
    soup.body.append(ol)

    return soup


def convert_opml_to_html(opml_path: Path, html_path: Path) -> None:
    logger.info("Reading OPML file: {}", opml_path)
    tree = ET.parse(opml_path)
    root = tree.getroot()

    body = root.find("body")
    if body is None:
        raise ValueError("No <body> found in OPML.")

    task_hierarchy = parse_opml_body(body)
    logger.info("Extracted {} top-level tasks", len(task_hierarchy))

    html_soup = build_html(task_hierarchy)
    logger.info("Writing HTML to {}", html_path)
    html_path.write_text(html_soup.prettify(), encoding="utf-8")


# --------------------------------------------------------------------------- #
# GUI
# --------------------------------------------------------------------------- #
class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("OneNote ↔ MLO Converter")
        self.geometry("560x260")
        self.resizable(False, False)

        pad = dict(padx=10, pady=8)

        # ---- Direction -------------------------------------------------------
        ttk.Label(self, text="Direction:").grid(row=0, column=0, sticky="w", **pad)
        self.direction_var = tk.StringVar(value="OneNote to MLO")
        ttk.Radiobutton(self, text="OneNote (.mht) → MLO (.opml)", variable=self.direction_var, value="OneNote to MLO", command=self.update_labels).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(self, text="MLO (.opml) → OneNote (.html)", variable=self.direction_var, value="MLO to OneNote", command=self.update_labels).grid(row=1, column=1, sticky="w")

        # ---- Input -------------------------------------------------------
        self.lbl_input = ttk.Label(self, text="Input .mht file:")
        self.lbl_input.grid(row=2, column=0, sticky="w", **pad)
        self.ent_input = ttk.Entry(self, width=50)
        self.ent_input.grid(row=2, column=1, **pad)
        self.btn_browse_input = ttk.Button(self, text="Browse…", command=self.browse_input)
        self.btn_browse_input.grid(row=2, column=2, **pad)

        # ---- Output ------------------------------------------------------
        self.lbl_output = ttk.Label(self, text="Output .opml file:")
        self.lbl_output.grid(row=3, column=0, sticky="w", **pad)
        self.ent_output = ttk.Entry(self, width=50)
        self.ent_output.grid(row=3, column=1, **pad)
        self.btn_browse_output = ttk.Button(self, text="Save as…", command=self.browse_output)
        self.btn_browse_output.grid(row=3, column=2, **pad)

        # ---- Convert -----------------------------------------------------
        ttk.Button(self, text="Convert →", command=self.run_convert).grid(row=4, column=1, pady=20)

        # auto-fill output when input changes
        self.ent_input.bind("<KeyRelease>", self.auto_output)

    # ------------------------------------------------------------------- #
    def update_labels(self) -> None:
        direction = self.direction_var.get()
        if direction == "OneNote to MLO":
            self.lbl_input.config(text="Input .mht file:")
            self.lbl_output.config(text="Output .opml file:")
        else:
            self.lbl_input.config(text="Input .opml file:")
            self.lbl_output.config(text="Output .html file:")
        self.auto_output()

    def browse_input(self) -> None:
        direction = self.direction_var.get()
        if direction == "OneNote to MLO":
            filetypes = [("MHTML files", "*.mht *.mhtml"), ("All files", "*.*")]
            title = "Select OneNote .mht file"
        else:
            filetypes = [("OPML files", "*.opml"), ("XML files", "*.xml"), ("All files", "*.*")]
            title = "Select MLO .opml file"

        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            self.ent_input.delete(0, tk.END)
            self.ent_input.insert(0, path)
            self.auto_output()

    def browse_output(self) -> None:
        direction = self.direction_var.get()
        if direction == "OneNote to MLO":
            defaultext = ".opml"
            filetypes = [("OPML files", "*.opml"), ("XML files", "*.xml")]
            title = "Save OPML file"
        else:
            defaultext = ".html"
            filetypes = [("HTML files", "*.html *.htm")]
            title = "Save HTML file"

        path = filedialog.asksaveasfilename(title=title, defaultextension=defaultext, filetypes=filetypes)
        if path:
            self.ent_output.delete(0, tk.END)
            self.ent_output.insert(0, path)

    def auto_output(self, *_) -> None:
        inp = Path(self.ent_input.get())
        if not inp.exists():
            return

        direction = self.direction_var.get()
        if direction == "OneNote to MLO" and inp.suffix.lower() in {".mht", ".mhtml"}:
            out = inp.with_suffix(".opml")
        elif direction == "MLO to OneNote" and inp.suffix.lower() == ".opml":
            out = inp.with_suffix(".html")
        else:
            return

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

        direction = self.direction_var.get()
        try:
            if direction == "OneNote to MLO":
                convert_mht_to_opml(in_path, out_path)
                msg = f"OPML created!\n{out_path}\n\nImport into MLO → File → Import → OPML"
            else:
                convert_opml_to_html(in_path, out_path)
                msg = f"HTML created!\n{out_path}\n\nOpen in browser → Select All → Copy → Paste into OneNote"

            messagebox.showinfo("Success", msg)
        except Exception as exc:
            logger.exception("Conversion failed")
            messagebox.showerror("Conversion failed", str(exc))


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    logger.remove()
    logger.add(sys.stderr, level="INFO")
    app = ConverterApp()
    app.mainloop()
