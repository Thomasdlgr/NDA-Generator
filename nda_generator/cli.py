from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from nda_generator.pipeline import DEFAULT_MODEL, run_review

DEFAULT_NDA = "NDA_Exemple.docx"
DEFAULT_PLAYBOOK = "NDA_Playbook.xlsx"
DEFAULT_OUT = "NDA_revu.docx"


def main(argv: list[str] | None = None) -> int:
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
    parser.add_argument(
        "--strict-ops",
        action="store_true",
        help="Refuser le batch si delete + insert_* sur le même paragraphe (forcer replace)",
    )
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

    ok = run_review(
        nda_path=nda_path,
        playbook_path=playbook_path,
        out_path=out_path,
        author=args.author,
        model=args.model,
        strict_ops=args.strict_ops,
        log=log,
    )
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
