"""Synthèse HTML des modifications appliquées par issue (second appel Anthropic)."""

from __future__ import annotations

import html
import json
import logging
import re
from typing import Any

import nh3
from anthropic import Anthropic

from docx_editor import EditOperation

_SUMMARY_TAGS = frozenset({"p", "ul", "ol", "li", "strong", "em", "br"})

SUMMARY_SYSTEM_PROMPT = """Tu es un juriste senior. On te donne le nom d'une issue de playbook et la liste des \
modifications effectivement appliquées au NDA (JSON : actions docx_editor sur des paragraphes référencés).

Rédige un **compte rendu en français** pour un lecteur business ou juriste :
- Résume l'enjeu de l'issue et ce qui a été modifié dans le contrat (substance), sans jargon technique JSON.
- Explique le sens des ajustements (ex. durée, périmètre, niveau d'obligation).
- Reste factuel par rapport aux opérations listées.

**Format de sortie** : fragment HTML uniquement (pas de <!DOCTYPE>, html, head, body, script, style).
**Balises autorisées** : <p>, <ul>, <ol>, <li>, <strong>, <em>, <br>.
Au plus une liste à puces si plusieurs changements distincts. Environ 120 à 220 mots maximum."""


def _truncate(s: str | None, max_len: int = 600) -> str:
    if not s:
        return ""
    t = str(s)
    if len(t) <= max_len:
        return t
    return t[: max_len - 1] + "…"


def operations_to_summary_json(operations: list[EditOperation]) -> str:
    rows: list[dict[str, Any]] = []
    for op in operations:
        row: dict[str, Any] = {
            "action": op.action,
            "paragraph": op.paragraph,
            "occurrence": op.occurrence,
        }
        if op.action == "replace":
            row["find"] = _truncate(op.find)
            row["replace_with"] = _truncate(op.replace_with)
        elif op.action == "delete":
            row["text"] = _truncate(op.text)
        elif op.action in ("insert_after", "insert_before"):
            row["anchor"] = _truncate(op.anchor)
            row["text"] = _truncate(op.text)
        rows.append(row)
    return json.dumps(rows, ensure_ascii=False, indent=2)


def _strip_html_fence(text: str) -> str:
    text = text.strip()
    m = re.match(r"^```(?:html)?\s*([\s\S]*?)```\s*$", text)
    if m:
        return m.group(1).strip()
    return text


def _clean_summary_body(raw: str) -> str:
    text = _strip_html_fence(raw)
    if not text:
        return "<p><em>(Synthèse vide.)</em></p>"
    if "<" not in text:
        text = f"<p>{html.escape(text)}</p>"
    cleaned = nh3.clean(text, tags=_SUMMARY_TAGS, attributes={})
    if not cleaned.strip():
        return "<p><em>(Synthèse vide.)</em></p>"
    return cleaned


def format_report_article(issue_nom: str, body_inner_html: str) -> str:
    """Enveloppe compte rendu (body_inner déjà sûr ou issu de nh3 / html.escape)."""
    return (
        '<article class="report-issue">'
        f'<h4 class="report-issue-title">{html.escape(issue_nom)}</h4>'
        f'<div class="report-issue-body">{body_inner_html}</div>'
        "</article>"
    )


def format_static_issue_report(issue_nom: str, *plain_lines: str) -> str:
    """Paragraphes en texte brut (échappés)."""
    inner = "".join(f"<p>{html.escape(line)}</p>" for line in plain_lines if line)
    if not inner:
        inner = "<p><em>—</em></p>"
    return format_report_article(issue_nom, inner)


def summarize_applied_edits(
    client: Anthropic,
    model: str,
    issue_nom: str,
    operations: list[EditOperation],
    log: logging.Logger,
) -> str:
    """Appelle Claude pour résumer les opérations appliquées ; retourne un <article> HTML complet."""
    payload = operations_to_summary_json(operations)
    user_content = f"""## Issue (playbook)
{issue_nom}

## Modifications appliquées (JSON)
{payload}
"""
    try:
        msg = client.messages.create(
            model=model,
            max_tokens=4096,
            system=SUMMARY_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_content}],
        )
    except Exception:
        log.exception("Échec appel Anthropic pour la synthèse de l’issue « %s »", issue_nom[:80])
        return format_static_issue_report(
            issue_nom,
            "La synthèse automatique n’a pas pu être générée (erreur API).",
            "Les modifications Word ont bien été appliquées pour cette issue.",
        )

    text_blocks = [b.text for b in msg.content if b.type == "text"]
    raw = "\n".join(text_blocks)
    body = _clean_summary_body(raw)
    return format_report_article(issue_nom, body)
