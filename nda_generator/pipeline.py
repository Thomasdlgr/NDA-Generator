from __future__ import annotations

import logging
import os
from collections.abc import Callable
from pathlib import Path
from typing import Any

from anthropic import Anthropic

from docx_editor import Document

from nda_generator.document_context import build_paragraph_catalog
from nda_generator.issue_comments import add_issue_comments_for_operations
from nda_generator.llm_review import review_issue
from nda_generator.llm_summary import format_static_issue_report, summarize_applied_edits
from nda_generator.ops_logging import log_operations
from nda_generator.operations_validate import explain_delete_insert_violation
from nda_generator.playbook import load_playbook

DEFAULT_MODEL = "claude-sonnet-4-20250514"

ProgressCallback = Callable[[dict[str, Any]], None]


def run_review(
    *,
    nda_path: Path,
    playbook_path: Path,
    out_path: Path,
    author: str,
    model: str | None = None,
    strict_ops: bool = False,
    log: logging.Logger,
    on_progress: ProgressCallback | None = None,
) -> bool:
    """Exécute la revue playbook complète. Retourne True si le document a été enregistré."""
    nda_path = nda_path.resolve()
    playbook_path = playbook_path.resolve()
    out_path = out_path.resolve()
    model = model or os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        log.error("Variable d'environnement ANTHROPIC_API_KEY manquante.")
        return False

    if not nda_path.is_file():
        log.error("Fichier NDA introuvable : %s", nda_path)
        return False
    if not playbook_path.is_file():
        log.error("Playbook introuvable : %s", playbook_path)
        return False

    log.info("NDA=%s | playbook=%s | sortie=%s", nda_path, playbook_path, out_path)

    issues = load_playbook(playbook_path)
    if not issues:
        log.error("Aucune issue chargée depuis le playbook.")
        return False
    log.info("%d issues chargées depuis le playbook", len(issues))

    n_issues = len(issues)
    if on_progress:
        on_progress(
            {
                "kind": "init",
                "total": n_issues,
                "issues": [i.nom for i in issues],
                "percent": 0,
            }
        )

    client = Anthropic(api_key=api_key)
    author_name = (author or "").strip() or "Revue playbook"

    try:
        doc = Document.open(nda_path, author=author_name)
    except Exception as e:
        log.error("Impossible d'ouvrir le document : %s", e)
        return False

    def _pct_done(idx: int) -> int:
        if n_issues <= 0:
            return 100
        return min(100, int(round(100 * idx / n_issues)))

    try:
        for idx, issue in enumerate(issues, start=1):

            def emit_issue_end(status: str, summary_html: str) -> None:
                if on_progress:
                    on_progress(
                        {
                            "kind": "issue_end",
                            "index": idx,
                            "total": n_issues,
                            "title": issue.nom,
                            "status": status,
                            "percent": _pct_done(idx),
                            "summary_html": summary_html,
                        }
                    )

            if on_progress:
                on_progress(
                    {
                        "kind": "issue_begin",
                        "index": idx,
                        "total": n_issues,
                        "title": issue.nom,
                        "percent": _pct_done(idx - 1),
                        "preferred_position": issue.preferred_position,
                        "fallback_position": issue.fallback_position,
                        "preferred_wording": issue.preferred_wording,
                    }
                )

            end_status = "ok"
            catalog = build_paragraph_catalog(doc)
            log.info("Issue %d/%d : %s", idx, n_issues, issue.nom[:80])
            try:
                operations, llm_json_text = review_issue(
                    client,
                    model,
                    issue.nom,
                    issue.preferred_position,
                    issue.fallback_position,
                    issue.preferred_wording,
                    catalog,
                )
            except Exception:
                log.exception("Échec LLM pour l'issue « %s »", issue.nom)
                end_status = "llm_error"
                emit_issue_end(
                    end_status,
                    format_static_issue_report(
                        issue.nom,
                        "Erreur lors de l’appel au modèle de revue.",
                        "Aucune modification n’a été appliquée pour cette issue.",
                    ),
                )
                continue

            log.debug("Réponse JSON LLM (normalisée) :\n%s", llm_json_text)

            if not operations:
                log.info("Aucune opération proposée.")
                end_status = "no_ops"
                emit_issue_end(
                    end_status,
                    format_static_issue_report(
                        issue.nom,
                        "Le modèle n’a proposé aucune modification du NDA pour cette issue.",
                        "Aucun markup Word n’a été appliqué sur ce point.",
                    ),
                )
                continue

            log.info("%d opération(s) à appliquer — détail :", len(operations))
            log_operations(log, operations, prefix=issue.nom[:60])

            violation = explain_delete_insert_violation(operations)
            if violation:
                if strict_ops:
                    log.error("Lot rejeté (--strict-ops) pour « %s » : %s", issue.nom, violation)
                    end_status = "strict_rejected"
                    emit_issue_end(
                        end_status,
                        format_static_issue_report(
                            issue.nom,
                            f"Lot d’opérations rejeté (mode strict) : {violation}",
                            "Aucune modification n’a été appliquée pour cette issue.",
                        ),
                    )
                    continue
                log.warning("Motif d'alerte révisions : %s", violation)

            try:
                doc.batch_edit(operations)
            except Exception as e:
                log.error(
                    "batch_edit refusé pour « %s » (%s). Opérations ignorées pour cette issue.",
                    issue.nom,
                    e,
                )
                end_status = "batch_error"
                emit_issue_end(
                    end_status,
                    format_static_issue_report(
                        issue.nom,
                        "Les modifications n’ont pas pu être enregistrées dans le document Word.",
                        f"Détail technique : {e!s}",
                    ),
                )
                continue

            add_issue_comments_for_operations(doc, issue.nom, operations, log)

            summary_html = summarize_applied_edits(client, model, issue.nom, operations, log)
            emit_issue_end(end_status, summary_html)

        doc.save(out_path)
        log.info("Document enregistré : %s", out_path)
        log.info(
            "Astuce Word : si les suppressions semblent « en commentaire », vérifiez "
            "Révision > Affichage des marques (bulles vs tout en ligne)."
        )
    finally:
        doc.close(cleanup=True)

    return True
