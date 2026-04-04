"""HTML de prévisualisation d'un .docx avec révisions Word (ins/del) en surbrillance."""

from __future__ import annotations

import html
import io
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _local_tag(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag


def _run_text(r: ET.Element) -> str:
    parts: list[str] = []
    for c in r:
        ln = _local_tag(c.tag)
        if ln == "t":
            parts.append(c.text or "")
        elif ln == "tab":
            parts.append(" ")
        elif ln == "br":
            parts.append("\n")
        elif ln == "noBreakHyphen":
            parts.append("\u2011")
    return "".join(parts)


def _inline_child_segments(child: ET.Element) -> list[tuple[str, str]]:
    ln = _local_tag(child.tag)
    if ln == "r":
        t = _run_text(child)
        return [("normal", t)] if t else []
    if ln == "ins":
        t = "".join(child.itertext())
        return [("ins", t)] if t else []
    if ln == "del":
        t = "".join(child.itertext())
        return [("del", t)] if t else []
    if ln == "moveFrom":
        t = "".join(child.itertext())
        return [("del", t)] if t else []
    if ln == "moveTo":
        t = "".join(child.itertext())
        return [("ins", t)] if t else []
    if ln == "hyperlink":
        out: list[tuple[str, str]] = []
        for c in child:
            out.extend(_inline_child_segments(c))
        return out
    if ln == "sdt":
        inner = child.find(_q("sdtContent"))
        if inner is None:
            return []
        out = []
        for c in inner:
            out.extend(_inline_child_segments(c))
        return out
    if ln == "smartTag":
        out = []
        for c in child:
            out.extend(_inline_child_segments(c))
        return out
    if ln == "fldSimple":
        t = "".join(child.itertext())
        return [("normal", t)] if t else []
    if ln in ("bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd"):
        return []
    t = "".join(child.itertext())
    if t.strip():
        return [("normal", t)]
    return []


def _merge_adjacent(segs: list[tuple[str, str]]) -> list[tuple[str, str]]:
    if not segs:
        return []
    out: list[list[str]] = [[segs[0][0], segs[0][1]]]
    for kind, text in segs[1:]:
        if kind == out[-1][0]:
            out[-1][1] += text
        else:
            out.append([kind, text])
    return [(a, b) for a, b in out]


def _escaped_with_line_breaks(text: str) -> str:
    esc = html.escape(text.replace("\r\n", "\n").replace("\r", "\n"))
    return esc.replace("\n", "<br />")


def _segments_to_html(segs: list[tuple[str, str]]) -> str:
    parts: list[str] = []
    for kind, text in _merge_adjacent(segs):
        if not text:
            continue
        if kind == "normal":
            parts.append(_escaped_with_line_breaks(text))
        elif kind == "ins":
            parts.append(f'<span class="rev-ins">{_escaped_with_line_breaks(text)}</span>')
        elif kind == "del":
            parts.append(f'<span class="rev-del">{_escaped_with_line_breaks(text)}</span>')
    return "".join(parts) if parts else "&nbsp;"


def _render_paragraph(p: ET.Element) -> str:
    segs: list[tuple[str, str]] = []
    for child in p:
        if _local_tag(child.tag) == "pPr":
            continue
        segs.extend(_inline_child_segments(child))
    inner = _segments_to_html(segs)
    return f"<p>{inner}</p>"


def _render_tc(tc: ET.Element) -> str:
    parts: list[str] = []
    for el in tc:
        ln = _local_tag(el.tag)
        if ln == "tcPr":
            continue
        if ln == "p":
            parts.append(_render_paragraph(el))
        elif ln == "tbl":
            parts.append(_render_table(el))
        else:
            parts.append(_render_body_element(el))
    return "".join(parts) if parts else "&nbsp;"


def _render_tr(tr: ET.Element) -> str:
    cells: list[str] = []
    for el in tr:
        ln = _local_tag(el.tag)
        if ln == "trPr":
            continue
        if ln == "tc":
            cells.append(f"<td>{_render_tc(el)}</td>")
    return "<tr>" + "".join(cells) + "</tr>"


def _render_table(tbl: ET.Element) -> str:
    rows: list[str] = []
    for el in tbl:
        ln = _local_tag(el.tag)
        if ln in ("tblPr", "tblGrid"):
            continue
        if ln == "tr":
            rows.append(_render_tr(el))
    return "<table>" + "".join(rows) + "</table>"


def _render_body_element(el: ET.Element) -> str:
    ln = _local_tag(el.tag)
    if ln == "p":
        return _render_paragraph(el)
    if ln == "tbl":
        return _render_table(el)
    if ln == "sdt":
        inner = el.find(_q("sdtContent"))
        if inner is None:
            return ""
        return "".join(filter(None, (_render_body_element(c) for c in inner)))
    if ln in ("sectPr", "proofErr"):
        return ""
    return ""


def docx_revision_html_fragment(docx_bytes: bytes) -> str:
    """Extrait le corps du document et produit du HTML (insertions / suppressions marquées)."""
    buf = io.BytesIO(docx_bytes)
    with zipfile.ZipFile(buf, "r") as zf:
        try:
            xml_bytes = zf.read("word/document.xml")
        except KeyError as e:
            raise ValueError("word/document.xml manquant") from e

    root = ET.fromstring(xml_bytes)
    body = root.find(_q("body"))
    if body is None:
        return "<p>(Corps du document introuvable.)</p>"

    chunks: list[str] = []
    for child in body:
        h = _render_body_element(child)
        if h:
            chunks.append(h)

    return "\n".join(chunks) if chunks else "<p>(Document vide.)</p>"
