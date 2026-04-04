from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

from anthropic import Anthropic
from dotenv import load_dotenv

from docx_editor import Document

from nda_generator.document_context import build_paragraph_catalog
from nda_generator.llm_review import review_issue
from nda_generator.playbook import load_playbook

DEFAULT_MODEL = "claude-sonnet-4-20250514"

DEFAULT_NDA = "NDA_Exemple.docx"
DEFAULT_PLAYBOOK = "NDA_Playbook.xlsx"
DEFAULT_OUT = "NDA_revu.docx"


def main(argv: list[str] | None = None) -> int:
    # Charge .env depuis le répertoire courant ou un dossier parent (find_dotenv).
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Revue d'un NDA (DOCX) selon un playbook Excel ; sortie en révisions Word via docx_editor."
    )
    parser.add_argument(
        "--nda",
        type=Path,
        default=None,
        help=f"Chemin du NDA (.docx). Défaut : ./{DEFAULT_NDA} dans le répertoire courant",
    )
    parser.add_argument(
        "--playbook",
        type=Path,
        default=None,
        help=f"Chemin du playbook (.xlsx). Défaut : ./{DEFAULT_PLAYBOOK}",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help=f"DOCX de sortie. Défaut : ./{DEFAULT_OUT}",
    )
    parser.add_argument(
        "--author",
        default="Revue playbook (IA)",
        help="Auteur des modifications suivies dans Word",
    )
    parser.add_argument(
        "--model",
        default=os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL),
        help="Modèle Anthropic (défaut: variable ANTHROPIC_MODEL ou Sonnet 4)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Logs détaillés")
    args = parser.parse_args(argv)

    cwd = Path.cwd()
    nda_path = (args.nda or cwd / DEFAULT_NDA).resolve()
    playbook_path = (args.playbook or cwd / DEFAULT_PLAYBOOK).resolve()
    out_path = (args.out or cwd / DEFAULT_OUT).resolve()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(message)s",
    )
    log = logging.getLogger("nda_generator")

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        log.error("Variable d'environnement ANTHROPIC_API_KEY manquante.")
        return 1

    if not nda_path.is_file():
        log.error("Fichier NDA introuvable : %s (utilisez --nda ou lancez depuis le dossier du projet)", nda_path)
        return 1
    if not playbook_path.is_file():
        log.error(
            "Playbook introuvable : %s (utilisez --playbook ou lancez depuis le dossier du projet)",
            playbook_path,
        )
        return 1

    log.info("NDA=%s | playbook=%s | sortie=%s", nda_path, playbook_path, out_path)

    issues = load_playbook(playbook_path)
    if not issues:
        log.error("Aucune issue chargée depuis le playbook.")
        return 1
    log.info("%d issues chargées depuis le playbook", len(issues))

    client = Anthropic(api_key=api_key)

    doc = Document.open(nda_path, author=args.author)
    try:
        for idx, issue in enumerate(issues, start=1):
            catalog = build_paragraph_catalog(doc)
            log.info("Issue %d/%d : %s", idx, len(issues), issue.nom[:80])
            try:
                operations, commentaire = review_issue(
                    client,
                    args.model,
                    issue.nom,
                    issue.preferred_position,
                    issue.fallback_position,
                    issue.preferred_wording,
                    catalog,
                )
            except Exception as e:
                log.exception("Échec LLM pour l'issue « %s » : %s", issue.nom, e)
                continue

            if commentaire:
                log.info("Synthèse : %s", commentaire)
            if not operations:
                log.info("Aucune opération proposée.")
                continue

            log.info("%d opération(s) à appliquer", len(operations))
            try:
                doc.batch_edit(operations)
            except Exception as e:
                log.error(
                    "batch_edit refusé pour « %s » (%s). Opérations ignorées pour cette issue.",
                    issue.nom,
                    e,
                )
                continue

        doc.save(out_path)
        log.info("Document enregistré : %s", out_path)
    finally:
        doc.close(cleanup=True)

    return 0


if __name__ == "__main__":
    sys.exit(main())
