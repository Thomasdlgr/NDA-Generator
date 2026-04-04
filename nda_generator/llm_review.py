from __future__ import annotations

import json
import re
from typing import Any

from anthropic import Anthropic

from docx_editor import EditOperation

SYSTEM_PROMPT = """Tu es un juriste senior chargé de comparer un NDA (texte fourni par paragraphe) \
avec une issue de playbook de négociation.

Tu dois proposer uniquement des modifications concrètes réalisables avec la librairie docx_editor, \
sous forme de liste d'objets JSON « EditOperation ».

Règles strictes :
- Utilise UNIQUEMENT des références de paragraphe exactement au format P{n}#{hash4} (ex. P14#8d4b), \
telles qu'indiquées dans le catalogue. Jamais de texte après un « | ».

Remplacement de texte (obligatoire pour toute substitution) :
- Dès que tu retires un passage pour le remplacer par une nouvelle rédaction, utilise **une seule** \
opération « replace » : find = sous-chaîne EXACTE du texte à retirer, replace_with = le nouveau texte.
- C'est ce qui produit dans Word le comportement attendu : l'ancien texte en **révision supprimée** \
(texte barré en suivi des modifications) et le nouveau en **révision insérée**, avec la **même mise en \
forme de caractères** que le passage remplacé (comportement natif de docx_editor).
- **Interdit** pour un même changement rédactionnel : enchaîner « delete » puis « insert_after » ou \
« insert_before » (risque d'ancres invalides, révisions non appariées, mise en forme incohérente).

Autres actions :
- « delete » seul : uniquement pour supprimer du texte **sans** le remplacer par du neuf au même endroit.
- « insert_after » / « insert_before » : uniquement pour **ajouter** du contenu (sans retirer en parallèle \
un fragment que tu remplaces par ce nouvel ajout).
- « find », « text », « anchor » : sous-chaînes EXACTES du paragraphe (espaces, ponctuation, guillemets \
« » vs " comme dans le source).
- « occurrence » : 0 = première occurrence dans le paragraphe, 1 = deuxième, etc.
- Préfère peu d'opérations ciblées. Si le NDA est déjà aligné avec le playbook, renvoie une liste vide.
- Ne invente pas de citations : si tu ne peux pas citer un extrait exact du catalogue, n'ajoute pas d'opération.

Schéma d'une opération :
- replace : action, paragraph, find, replace_with, occurrence (optionnel, défaut 0)
- delete : action, paragraph, text, occurrence (optionnel)
- insert_after / insert_before : action, paragraph, anchor, text, occurrence (optionnel)

Réponds UNIQUEMENT avec un objet JSON : {"operations":[...], "commentaire":"court résumé en français"} \
sans markdown ni texte autour."""


def _extract_json_object(text: str) -> dict[str, Any]:
    text = text.strip()
    fence = re.match(r"^```(?:json)?\s*([\s\S]*?)```\s*$", text)
    if fence:
        text = fence.group(1).strip()
    return json.loads(text)


def _to_operation(raw: dict[str, Any]) -> EditOperation:
    action = raw["action"]
    if action not in ("replace", "delete", "insert_after", "insert_before"):
        raise ValueError(f"action invalide: {action}")
    paragraph = str(raw["paragraph"]).split("|")[0].strip()
    occurrence = int(raw.get("occurrence", 0))
    if action == "replace":
        return EditOperation(
            action="replace",
            paragraph=paragraph,
            find=str(raw["find"]),
            replace_with=str(raw["replace_with"]),
            occurrence=occurrence,
        )
    if action == "delete":
        return EditOperation(
            action="delete",
            paragraph=paragraph,
            text=str(raw["text"]),
            occurrence=occurrence,
        )
    if action in ("insert_after", "insert_before"):
        return EditOperation(
            action=action,  # type: ignore[arg-type]
            paragraph=paragraph,
            anchor=str(raw["anchor"]),
            text=str(raw["text"]),
            occurrence=occurrence,
        )
    raise ValueError(f"action non gérée: {action}")


def review_issue(
    client: Anthropic,
    model: str,
    issue_nom: str,
    preferred: str,
    fallback: str,
    preferred_wording: str,
    paragraph_catalog: str,
) -> tuple[list[EditOperation], str]:
    """Appelle Claude et retourne (EditOperation[], commentaire)."""
    user_content = f"""## Issue (playbook)
**Nom :** {issue_nom}

**Position préférée :**
{preferred}

**Fallback :**
{fallback}

**Libellé préféré (cible rédactionnelle) :**
{preferred_wording}

## Catalogue du NDA (référence de paragraphe + texte intégral du paragraphe)
Les blocs sont séparés par ---. Chaque bloc commence par la référence P{{n}}#{{hash}} sur sa propre ligne.

{paragraph_catalog}
"""

    msg = client.messages.create(
        model=model,
        max_tokens=8192,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )
    text_blocks = [b.text for b in msg.content if b.type == "text"]
    raw_text = "\n".join(text_blocks)
    data = _extract_json_object(raw_text)
    ops_raw = data.get("operations") or []
    if not isinstance(ops_raw, list):
        raise ValueError("Le JSON doit contenir une clé 'operations' (liste)")
    operations = [_to_operation(item) for item in ops_raw]
    commentaire = str(data.get("commentaire", "")).strip()
    return operations, commentaire
