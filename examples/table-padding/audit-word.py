"""Audit a Word Save As copy, after altChunk has become native table content.

Usage: python3 audit-word.py /path/to/padding-native.docx
The export itself must be opened and saved as a separate copy in Word first.
"""

import json
import sys
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": NAMESPACE}
W = "{" + NAMESPACE + "}"
source = Path(sys.argv[1])
with ZipFile(source) as archive:
    document = ET.fromstring(archive.read("word/document.xml"))

assert document.find(".//w:altChunk", NS) is None, "Save a native copy in Word first."
cells = []
for cell in document.findall(".//w:tc", NS):
    label = "".join(node.text or "" for node in cell.findall(".//w:t", NS))
    if not label.startswith(("BG0", "TOP", "MIDDLE", "BOTTOM")):
        continue
    margins = cell.find("w:tcPr/w:tcMar", NS)
    padding = {} if margins is None else {
        node.tag.removeprefix(W): int(node.get(W + "w")) for node in margins
    }
    assert padding == {"top": 80, "left": 80, "bottom": 80, "right": 80}, label
    alignment = cell.find("w:tcPr/w:vAlign", NS)
    valign = "top" if alignment is None else alignment.get(W + "val")
    expected = "top" if label.startswith("TOP") else "bottom" if label.startswith("BOTTOM") else "center"
    assert valign == expected, (label, valign)
    spacing = cell.find("w:p/w:pPr/w:spacing", NS)
    after = None if spacing is None else spacing.get(W + "after")
    assert after == "0", (label, "Paragraph after-spacing must explicitly override Word defaults.")
    cells.append({"label": label, "paddingTwips": padding, "verticalAlign": valign, "paragraphAfterTwips": 0})

assert len(cells) == 7, f"Expected 7 diagnostic cells, found {len(cells)}."
print(json.dumps({"nativeDocument": source.name, "cells": cells, "status": "passed"}, ensure_ascii=False, indent=2))
