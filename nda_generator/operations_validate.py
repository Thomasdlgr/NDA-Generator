from __future__ import annotations

from collections import defaultdict

from docx_editor import EditOperation


def paragraph_key(op: EditOperation) -> str:
    return (op.paragraph or "").split("|")[0].strip()


def find_delete_plus_insert_same_paragraph(operations: list[EditOperation]) -> list[str]:
    """Retourne la liste des refs de paragraphe où coexistent delete et insert_* dans le même batch."""
    flags: dict[str, dict[str, bool]] = defaultdict(lambda: {"delete": False, "insert": False})
    for op in operations:
        key = paragraph_key(op)
        if not key:
            continue
        if op.action == "delete":
            flags[key]["delete"] = True
        elif op.action in ("insert_after", "insert_before"):
            flags[key]["insert"] = True
    return [p for p, f in flags.items() if f["delete"] and f["insert"]]


def explain_delete_insert_violation(operations: list[EditOperation]) -> str | None:
    bad = find_delete_plus_insert_same_paragraph(operations)
    if not bad:
        return None
    return (
        "Même batch : delete + insert_after/insert_before sur le(s) paragraphe(s) "
        + ", ".join(bad)
        + ". Préférer une seule opération « replace » pour un remplacement (révision barrée + insérée)."
    )
