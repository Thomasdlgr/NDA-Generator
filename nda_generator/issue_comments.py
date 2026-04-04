from __future__ import annotations

import logging

from docx_editor import Document, EditOperation
from docx_editor.exceptions import CommentError, TextNotFoundError


def _non_empty(s: str | None) -> str | None:
    if s is None:
        return None
    t = str(s).strip()
    return t or None


def _anchor_candidates(op: EditOperation) -> list[str]:
    """Sous-chaînes à tenter pour add_comment (texte présent dans un w:t après l'édition)."""
    out: list[str] = []
    if op.action == "replace":
        for key in (_non_empty(op.replace_with), _non_empty(op.find)):
            if key and key not in out:
                out.append(key)
    elif op.action in ("insert_after", "insert_before"):
        key = _non_empty(op.text)
        if key:
            out.append(key)
    elif op.action == "delete":
        key = _non_empty(op.text)
        if key:
            out.append(key)
    return out


def add_issue_comments_for_operations(
    doc: Document,
    issue_nom: str,
    operations: list[EditOperation],
    log: logging.Logger,
) -> None:
    """Ajoute un commentaire Word par opération, texte du commentaire = nom de l'issue playbook."""
    if not operations or not issue_nom.strip():
        return
    comment_text = issue_nom.strip()
    for i, op in enumerate(operations):
        placed = False
        for anchor in _anchor_candidates(op):
            try:
                doc.add_comment(anchor, comment_text)
                placed = True
                log.debug("Commentaire playbook ancré sur %r (op #%d %s)", anchor[:80], i, op.action)
                break
            except (TextNotFoundError, CommentError):
                continue
        if not placed:
            log.warning(
                "Impossible d'ancrer un commentaire pour l'issue « %s », op #%d (%s). "
                "Ancres essayées : %s — (souvent le cas d'une suppression pure, texte dans w:delText).",
                issue_nom[:70],
                i,
                op.action,
                _anchor_candidates(op),
            )
