#!/usr/bin/env python
"""
Convert an OPML file from MyLifeOrganized (MLO) into an HTML format suitable for import into OneNote.

Generates a numbered list with:
  - Font: Calibri, 11pt
  - Numbering: 1. → a. → i. → 1) → 1. (repeating)
"""

import sys
from pathlib import Path
from typing import List

import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup, Tag
from loguru import logger


# --------------------------------------------------------------------------- #
# Helper: parse OPML tree → list[dict]
# --------------------------------------------------------------------------- #
def parse_opml_outline(outline_elem: ET.Element) -> dict[str, any]:
    """
    Recursively parse an <outline> element into a task dictionary.

    :param outline_elem: The <outline> ET.Element to parse.
    :return: Dictionary with 'name' and 'subtasks' (list of similar dicts).
    """
    name = outline_elem.get("text", "(no text)")
    subtasks = [parse_opml_outline(sub) for sub in outline_elem]
    return {"name": name, "subtasks": subtasks}


def parse_opml_body(body_elem: ET.Element) -> List[dict[str, any]]:
    """
    Parse the <body> of the OPML into a list of top-level task dictionaries.

    :param body_elem: The <body> ET.Element.
    :return: List of task dicts from top-level <outline> elements.
    """
    return [parse_opml_outline(outline) for outline in body_elem.findall("outline")]


# --------------------------------------------------------------------------- #
# Helper: build HTML <ol><li> with custom numbering
# --------------------------------------------------------------------------- #
NUMBERING_STYLES = [
    "decimal",        # 1., 2., 3.
    "lower-alpha",    # a., b., c.
    "lower-roman",    # i., ii., iii.
    "decimal)",       # 1), 2), 3)
    "decimal",        # 1., 2., 3. (repeat)
]

def build_html_li(task: dict[str, any], soup: BeautifulSoup, level: int = 0) -> Tag:
    """
    Recursively build an <li> element with nested <ol> using custom numbering.

    :param task: Task dictionary {'name': str, 'subtasks': [...]}
    :param soup: BeautifulSoup instance for creating tags.
    :param level: Current nesting level (0 = top).
    :return: The <li> Tag.
    """
    li = soup.new_tag("li")
    span = soup.new_tag("span", style="font-family:Calibri;font-size:11pt")
    span.string = task["name"]
    li.append(span)

    if task["subtasks"]:
        style_idx = level % len(NUMBERING_STYLES)
        list_style = NUMBERING_STYLES[style_idx]
        ol = soup.new_tag("ol", style=f"list-style-type:{list_style};margin-left:36pt")
        for sub in task["subtasks"]:
            ol.append(build_html_li(sub, soup, level + 1))
        li.append(ol)

    return li


def build_html(task_list: List[dict[str, any]]) -> BeautifulSoup:
    """
    Build a complete HTML document with the task hierarchy.

    :param task_list: List of top-level task dicts.
    :return: BeautifulSoup object ready to write.
    """
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

    # Start with level 0 (1., 2., 3.)
    ol = soup.new_tag("ol", style="list-style-type:decimal;margin-left:36pt")
    for task in task_list:
        ol.append(build_html_li(task, soup, level=0))
    soup.body.append(ol)

    return soup


# --------------------------------------------------------------------------- #
# Main conversion
# --------------------------------------------------------------------------- #
def convert_opml_to_html(input_path: Path, output_path: Path) -> None:
    """
    Convert MLO OPML to HTML for OneNote import.

    :param input_path: Path to .opml input.
    :param output_path: Path to .html output.
    :raises ValueError: If invalid OPML structure.
    """
    logger.info("Reading OPML file: {}", input_path)
    tree = ET.parse(input_path)
    root = tree.getroot()

    body = root.find("body")
    if body is None:
        raise ValueError("No <body> found in OPML.")

    task_hierarchy = parse_opml_body(body)
    logger.info("Extracted {} top-level tasks", len(task_hierarchy))

    html_soup = build_html(task_hierarchy)
    logger.info("Writing HTML to {}", output_path)
    output_path.write_text(html_soup.prettify(), encoding="utf-8")


# --------------------------------------------------------------------------- #
# CLI entry-point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python mlo_to_onenote.py <input.opml> <output.html>")
        sys.exit(1)

    in_file = Path(sys.argv[1])
    out_file = Path(sys.argv[2])

    try:
        convert_opml_to_html(in_file, out_file)
        print(f"Success: {out_file} created")
        print("Open in browser → Select All → Copy → Paste into OneNote")
    except Exception as e:
        logger.error("Error: {}", e)
        print(f"Error: {e}")
        sys.exit(1)