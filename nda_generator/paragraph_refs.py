"""Extraction des indices de paragraphe P{n}#… à partir des EditOperation."""

from __future__ import annotations

import re

from docx_editor import EditOperation

_P_REF = re.compile(r"^P(\d+)#", re.IGNORECASE)


def paragraph_indices_from_operations(operations: list[EditOperation]) -> list[int]:
    """Indices 1-based alignés sur `Document.list_paragraphs` / prévisualisation HTML."""
    seen: set[int] = set()
    for op in operations:
        m = _P_REF.match((op.paragraph or "").strip())
        if m:
            seen.add(int(m.group(1)))
    return sorted(seen)
