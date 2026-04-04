from __future__ import annotations

from docx_editor import Document


def build_paragraph_catalog(doc: Document) -> str:
    """Concatène chaque ref P{n}#hash avec le texte visible du paragraphe (ordre document)."""
    previews = doc.list_paragraphs(max_chars=0)
    bodies = doc.get_visible_text().split("\n")
    chunks: list[str] = []
    for i, prev in enumerate(previews):
        ref = prev.split("|")[0].strip()
        body = bodies[i] if i < len(bodies) else ""
        chunks.append(f"{ref}\n{body}")
    return "\n\n---\n\n".join(chunks)
