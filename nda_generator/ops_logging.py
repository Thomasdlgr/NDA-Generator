from __future__ import annotations

import logging

from docx_editor import EditOperation


def log_operations(logger: logging.Logger, operations: list[EditOperation], *, prefix: str = "") -> None:
    """Log chaque EditOperation avec tous les champs (repr pour textes multi-lignes)."""
    tag = f"{prefix} " if prefix else ""
    if not operations:
        logger.info("%sAucune opération dans ce lot.", tag)
        return
    for i, op in enumerate(operations):
        if op.action == "replace":
            logger.info(
                "%s#%d replace paragraph=%s occurrence=%d find=%r replace_with=%r",
                tag,
                i,
                op.paragraph,
                op.occurrence,
                op.find,
                op.replace_with,
            )
        elif op.action == "delete":
            logger.info(
                "%s#%d delete paragraph=%s occurrence=%d text=%r",
                tag,
                i,
                op.paragraph,
                op.occurrence,
                op.text,
            )
        elif op.action in ("insert_after", "insert_before"):
            logger.info(
                "%s#%d %s paragraph=%s occurrence=%d anchor=%r text=%r",
                tag,
                i,
                op.action,
                op.paragraph,
                op.occurrence,
                op.anchor,
                op.text,
            )
        else:
            logger.info("%s#%d %r", tag, i, op)
