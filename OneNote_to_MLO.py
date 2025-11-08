#!/usr/bin/env python
"""
Convert a OneNote page (exported as .mht) that contains a numbered list
into the OPML format for import into MyLifeOrganized (MLO).

Preserves the indentation hierarchy as nested <outline> elements.
MLO supports OPML import, which is simpler than their custom XML.
"""

import sys
import re
from pathlib import Path
from typing import Any, Dict, List

import quopri
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup, Tag, NavigableString
from loguru import logger


# --------------------------------------------------------------------------- #
# Helper: extract real HTML from .mht (handles MIME, quoted-printable)
# --------------------------------------------------------------------------- #
def extract_html_from_mht(input_path: Path) -> str:
    """
    Extract and decode the HTML content from a .mht file.

    Handles both single-part and multipart MIME, and decodes quoted-printable if needed.

    :param input_path: Path to the .mht file.
    :return: The decoded HTML string.
    """
    logger.info("Reading MHT file: {}", input_path)
    raw = input_path.read_text(encoding="utf-8", errors="replace")

    # Check for multipart boundary
    boundary_match = re.search(r'boundary\s*=\s*["\']?([^"\']*)["\']?', raw, re.I | re.M)
    if boundary_match:
        boundary = boundary_match.group(1)
        parts = re.split(rf"--{re.escape(boundary)}", raw)
        for part in parts[1:-1]:  # Skip empty first/last
            part = part.strip()
            if not part:
                continue
            header_body = re.split(r"\r?\n\r?\n|\n\n", part, maxsplit=1)
            if len(header_body) < 2:
                continue
            part_header, part_body = header_body
            if "text/html" in part_header.lower():
                if "quoted-printable" in part_header.lower():
                    logger.debug("Decoding quoted-printable content")
                    part_body = quopri.decodestring(part_body.encode("utf-8")).decode("utf-8")
                return part_body

    # Single part fallback
    header_body = re.split(r"\r?\n\r?\n|\n\n", raw, maxsplit=1)
    if len(header_body) < 2:
        return raw
    part_header, part_body = header_body
    if "quoted-printable" in part_header.lower():
        logger.debug("Decoding quoted-printable content")
        part_body = quopri.decodestring(part_body.encode("utf-8")).decode("utf-8")
    return part_body


# --------------------------------------------------------------------------- #
# Helper: clean text – remove &#10; and normalize whitespace
# --------------------------------------------------------------------------- #
def clean_text(text: str) -> str:
    """
    Remove HTML entities like &#10; and collapse whitespace.
    Preserves meaningful line breaks as single spaces.
    """
    # Decode &#10; (line feed) and other numeric entities
    text = re.sub(r"&#\d+;", " ", text)
    # Collapse multiple whitespace (including newlines) into single space
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# --------------------------------------------------------------------------- #
# Helper: parse <ol>/<ul> → list[dict] (handles non-standard nesting)
# --------------------------------------------------------------------------- #
def parse_list(list_tag: Tag) -> List[Dict[str, Any]]:
    """
    Recursively parse an <ol> or <ul> into a list of task dictionaries.

    Handles cases where sub-lists are siblings (not children) of the parent <li>,
    by assigning them to the most recent task.

    :param list_tag: BeautifulSoup <ol> or <ul> element.
    :return: List of {'name': str, 'subtasks': [...]}
    """
    if list_tag.name not in {"ol", "ul"}:
        raise ValueError(f"Expected <ol> or <ul>, got <{list_tag.name}>")

    tasks: List[Dict[str, Any]] = []
    current_task: Dict[str, Any] | None = None

    for child in list_tag.children:
        if isinstance(child, NavigableString):
            continue  # Skip whitespace/text nodes
        if child.name == "li":
            raw_text = child.get_text(separator=" ", strip=True)
            name = clean_text(raw_text)
            current_task = {"name": name, "subtasks": []}
            tasks.append(current_task)
            logger.debug("Added task: {}", name)
        elif child.name in {"ol", "ul"} and current_task:
            subtasks = parse_list(child)
            current_task["subtasks"].extend(subtasks)
            logger.debug("Added {} subtasks to '{}'", len(subtasks), current_task["name"])

    return tasks


# --------------------------------------------------------------------------- #
# Helper: build OPML tree
# --------------------------------------------------------------------------- #
def build_outline(task: Dict[str, Any]) -> ET.Element:
    """
    Recursively build an <outline> element.

    :param task: {'name': str, 'subtasks': [...]}
    :return: ET.Element for the outline.
    """
    elem = ET.Element("outline")
    elem.set("text", task["name"])

    for sub in task["subtasks"]:
        elem.append(build_outline(sub))

    return elem


def build_opml(task_list: List[Dict[str, Any]]) -> ET.ElementTree:
    """
    Build a complete OPML tree.

    :param task_list: Output from parse_list.
    :return: ET.ElementTree ready to write.
    """
    opml = ET.Element("opml")
    opml.set("version", "1.0")

    head = ET.SubElement(opml, "head")
    ET.SubElement(head, "title").text = "OneNote ToDo List"

    body = ET.SubElement(opml, "body")
    for t in task_list:
        body.append(build_outline(t))

    return ET.ElementTree(opml)


# --------------------------------------------------------------------------- #
# Main conversion
# --------------------------------------------------------------------------- #
def convert_mht_to_opml(input_path: Path, output_path: Path) -> None:
    """
    Convert OneNote .mht to OPML for MLO import.

    :param input_path: Path to .mht.
    :param output_path: Path to output OPML (.opml or .xml).
    :raises ValueError: If no list found or invalid structure.
    """
    html = extract_html_from_mht(input_path)
    logger.info("Parsing HTML content")
    soup = BeautifulSoup(html, "lxml")

    body = soup.body
    if not body:
        raise ValueError("No <body> tag found in the HTML.")

    main_list = body.find("ol") or body.find("ul")
    if not main_list:
        raise ValueError("No ordered or unordered list found in the body.")

    logger.info("Found main list: <{}>", main_list.name)
    task_hierarchy = parse_list(main_list)
    logger.info("Extracted {} top-level tasks", len(task_hierarchy))

    tree = build_opml(task_hierarchy)
    ET.indent(tree, space="  ", level=0)
    logger.info("Writing OPML to {}", output_path)
    tree.write(
        output_path,
        encoding="utf-8",
        xml_declaration=True,
        method="xml",
        short_empty_elements=False,
    )


# --------------------------------------------------------------------------- #
# CLI entry-point
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python onenote_to_mlo.py <input.mht> <output.opml>")
        sys.exit(1)

    in_file = Path(sys.argv[1])
    out_file = Path(sys.argv[2])

    try:
        convert_mht_to_opml(in_file, out_file)
        print(f"Success: {out_file} created – import into MLO as OPML")
    except Exception as e:
        logger.error("Error: {}", e)
        print(f"Error: {e}")
        sys.exit(1)
