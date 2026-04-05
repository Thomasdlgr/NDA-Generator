"""HTML de prévisualisation d'un .docx avec révisions Word (ins/del) et titres (h1–h6)."""

from __future__ import annotations

import html
import io
import re
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _local_tag(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag


def _clip_heading_level(n: int) -> int | None:
    if 1 <= n <= 6:
        return n
    if 7 <= n <= 9:
        return 6
    return None


def _display_name_to_heading_level(name: str) -> int | None:
    name = (name or "").strip()
    m = re.match(r"(?i)^heading\s*(\d)\s*$", name)
    if m:
        return _clip_heading_level(int(m.group(1)))
    m = re.match(r"(?i)^titre\s*(\d)\s*$", name)
    if m:
        return _clip_heading_level(int(m.group(1)))
    if re.match(r"(?i)^(title|titre)$", name):
        return 1
    if re.match(r"(?i)^(subtitle|sous-titre|soustitre)$", name):
        return 2
    return None


def _style_id_to_heading_level(style_id: str) -> int | None:
    sid = (style_id or "").strip()
    m = re.match(r"(?i)^heading\s*(\d)\s*$", sid)
    if m:
        return _clip_heading_level(int(m.group(1)))
    m = re.match(r"(?i)^titre(\d)$", sid)
    if m:
        return _clip_heading_level(int(m.group(1)))
    m = re.match(r"(?i)^heading(\d)$", sid)
    if m:
        return _clip_heading_level(int(m.group(1)))
    low = sid.lower()
    if low in ("title", "titre"):
        return 1
    if low in ("subtitle", "soustitre", "sous-titre"):
        return 2
    return None


def _parse_styles_heading_levels(xml_bytes: bytes) -> dict[str, int]:
    """Map w:styleId → niveau de titre (1–6) d’après w:name des styles de paragraphe."""
    levels: dict[str, int] = {}
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return levels
    for st in root.findall(_q("style")):
        if st.get(_q("type")) != "paragraph":
            continue
        sid = st.get(_q("styleId"))
        if not sid:
            continue
        name_el = st.find(_q("name"))
        if name_el is None:
            continue
        display = name_el.get(_q("val")) or ""
        lvl = _display_name_to_heading_level(display)
        if lvl is not None:
            levels[sid] = lvl
    return levels


def _paragraph_heading_level(p: ET.Element, style_levels: dict[str, int]) -> int | None:
    ppr = p.find(_q("pPr"))
    if ppr is None:
        return None
    style_id = None
    ps = ppr.find(_q("pStyle"))
    if ps is not None:
        style_id = ps.get(_q("val"))
    if style_id:
        if style_id in style_levels:
            return style_levels[style_id]
        hid = _style_id_to_heading_level(style_id)
        if hid is not None:
            return hid
    ol = ppr.find(_q("outlineLvl"))
    if ol is not None:
        raw = ol.get(_q("val"))
        if raw is not None and raw.isdigit():
            return _clip_heading_level(int(raw) + 1)
    return None


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


def _iter_wp_depth_first(el: ET.Element):
    """Ordre depth-first des w:p sous un élément (aligné sur getElementsByTagName('w:p') du corps)."""
    for child in el:
        if _local_tag(child.tag) == "p":
            yield child
        else:
            yield from _iter_wp_depth_first(child)


def _paragraph_index_map(body: ET.Element) -> dict[int, int]:
    return {id(p): i for i, p in enumerate(_iter_wp_depth_first(body), start=1)}


def _p_data_attr(p_el: ET.Element, idx_map: dict[int, int]) -> str:
    n = idx_map.get(id(p_el))
    if n is None:
        return ""
    return f' data-p-index="{n}"'


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


def _render_paragraph(p: ET.Element, style_levels: dict[str, int], idx_map: dict[int, int]) -> str:
    segs: list[tuple[str, str]] = []
    for child in p:
        if _local_tag(child.tag) == "pPr":
            continue
        segs.extend(_inline_child_segments(child))
    inner = _segments_to_html(segs)
    attr = _p_data_attr(p, idx_map)
    level = _paragraph_heading_level(p, style_levels)
    if level is not None:
        tag = f"h{level}"
        return f"<{tag}{attr}>{inner}</{tag}>"
    return f"<p{attr}>{inner}</p>"


def _render_tc(tc: ET.Element, style_levels: dict[str, int], idx_map: dict[int, int]) -> str:
    parts: list[str] = []
    for el in tc:
        ln = _local_tag(el.tag)
        if ln == "tcPr":
            continue
        if ln == "p":
            parts.append(_render_paragraph(el, style_levels, idx_map))
        elif ln == "tbl":
            parts.append(_render_table(el, style_levels, idx_map))
        else:
            parts.append(_render_body_element(el, style_levels, idx_map))
    return "".join(parts) if parts else "&nbsp;"


def _render_tr(tr: ET.Element, style_levels: dict[str, int], idx_map: dict[int, int]) -> str:
    cells: list[str] = []
    for el in tr:
        ln = _local_tag(el.tag)
        if ln == "trPr":
            continue
        if ln == "tc":
            cells.append(f"<td>{_render_tc(el, style_levels, idx_map)}</td>")
    return "<tr>" + "".join(cells) + "</tr>"


def _render_table(tbl: ET.Element, style_levels: dict[str, int], idx_map: dict[int, int]) -> str:
    rows: list[str] = []
    for el in tbl:
        ln = _local_tag(el.tag)
        if ln in ("tblPr", "tblGrid"):
            continue
        if ln == "tr":
            rows.append(_render_tr(el, style_levels, idx_map))
    return "<table>" + "".join(rows) + "</table>"


def _render_body_element(el: ET.Element, style_levels: dict[str, int], idx_map: dict[int, int]) -> str:
    ln = _local_tag(el.tag)
    if ln == "p":
        return _render_paragraph(el, style_levels, idx_map)
    if ln == "tbl":
        return _render_table(el, style_levels, idx_map)
    if ln == "sdt":
        inner = el.find(_q("sdtContent"))
        if inner is None:
            return ""
        return "".join(filter(None, (_render_body_element(c, style_levels, idx_map) for c in inner)))
    if ln in ("sectPr", "proofErr"):
        return ""
    return "".join(
        filter(None, (_render_body_element(c, style_levels, idx_map) for c in el))
    )


def docx_revision_html_fragment(docx_bytes: bytes) -> str:
    """Extrait le corps du document et produit du HTML (titres h1–h6, ins/del)."""
    buf = io.BytesIO(docx_bytes)
    with zipfile.ZipFile(buf, "r") as zf:
        try:
            xml_bytes = zf.read("word/document.xml")
        except KeyError as e:
            raise ValueError("word/document.xml manquant") from e
        try:
            styles_bytes = zf.read("word/styles.xml")
        except KeyError:
            styles_bytes = b""

    style_levels = _parse_styles_heading_levels(styles_bytes) if styles_bytes else {}

    root = ET.fromstring(xml_bytes)
    body = root.find(_q("body"))
    if body is None:
        return "<p>(Corps du document introuvable.)</p>"

    idx_map = _paragraph_index_map(body)

    chunks: list[str] = []
    for child in body:
        h = _render_body_element(child, style_levels, idx_map)
        if h:
            chunks.append(h)

    return "\n".join(chunks) if chunks else "<p>(Document vide.)</p>"
