#!/usr/bin/env python3
"""
PASSES 2–4 — Remplissage de contenu du rapport MOOMOO.

Idempotent lorsqu'appelé après normalisation (repart du doc déjà nettoyé en mémoire).
Ne modifie jamais le contenu de DÉDICACE, REMERCIEMENTS, PARTIE I.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

DATA_PATH = Path("/workspace/scripts/report_content_data.json")
REFS_BIBLIO = Path("/workspace/refs_biblio.json")
REFS_WEB = Path("/workspace/refs_web.json")
A_VERIFIER = Path("/workspace/A_VERIFIER.md")
PLACEHOLDERS = Path("/workspace/PLACEHOLDERS.md")

FONT_NAME = "Times New Roman"
FONT_BODY_PT = 12
FONT_CAPTION_PT = 10
LINE_TWIPS = 360
SPACING_AFTER_PT = 6

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}

PROTECTED_H1 = {
    "DÉDICACE",
    "DEDICACE",
    "REMERCIEMENTS",
    "PARTIE I : PHASE D'INSERTION",
    "PARTIE I : PHASE D’INSERTION",
}

PLACEHOLDER_RE = re.compile(
    r"^(Ins[eé]rer votre texte ici[.…\s]*|"
    r"Ins[eé]rez votre (texte|diagramme|figure|organigramme|tableau).*|"
    r"Inserez votre.*|"
    r"Insérez vos diagrammes ici[.…\s]*)$",
    re.I,
)

DIAGRAM_HINTS = re.compile(
    r"(diagramme|figure|gantt|capture|organigramme|exemple de)",
    re.I,
)


def _norm(s: str) -> str:
    s = s.replace("\xa0", " ").replace("’", "'").replace("‘", "'")
    s = re.sub(r"\s+", " ", s).strip()
    # Unify dossier numbering
    s = re.sub(r"^DOSSIER\s*IV\b", "DOSSIER 4", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*V\b", "DOSSIER 5", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*VI\b", "DOSSIER 6", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*VII\b", "DOSSIER 7", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*1\b", "DOSSIER 1", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*2\b", "DOSSIER 2", s, flags=re.I)
    s = re.sub(r"^DOSSIER\s*3\b", "DOSSIER 3", s, flags=re.I)
    s = re.sub(r"^PARTIE\s*2\b", "PARTIE 2", s, flags=re.I)
    s = re.sub(r":\s*$", "", s)
    return s


def _load_data() -> dict[str, Any]:
    return json.loads(DATA_PATH.read_text(encoding="utf-8"))


def _ensure_rfonts(rpr: etree._Element) -> None:
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{attr}"), FONT_NAME)


def _make_run(text: str, *, italic: bool = False, size_pt: float = FONT_BODY_PT) -> etree._Element:
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    _ensure_rfonts(rpr)
    half = str(int(size_pt * 2))
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), half)
        rpr.append(el)
    if italic:
        rpr.append(OxmlElement("w:i"))
        rpr.append(OxmlElement("w:iCs"))
    r.append(rpr)
    t = OxmlElement("w:t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _set_p_format(p: etree._Element, *, center: bool = False) -> None:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), "center" if center else "both")
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(LINE_TWIPS))
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:after"), str(int(SPACING_AFTER_PT * 20)))


def _clear_p(p: etree._Element) -> None:
    ppr = p.find(qn("w:pPr"))
    for child in list(p):
        if child is not ppr:
            p.remove(child)


def _rewrite_p(p: etree._Element, text: str, *, italic: bool = False) -> None:
    _clear_p(p)
    _set_p_format(p)
    p.append(_make_run(text, italic=italic))
    ppr = p.find(qn("w:pPr"))
    if ppr is not None and p[0] is not ppr:
        p.remove(ppr)
        p.insert(0, ppr)


def _para_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _style_id(p: etree._Element) -> str:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return ""
    ps = ppr.find(qn("w:pStyle"))
    return (ps.get(qn("w:val")) or "") if ps is not None else ""


def _is_heading(style_id: str, text: str) -> int:
    """Return heading level 1..4 or 0."""
    sid = style_id.lower()
    if sid in {"titre1", "heading1"}:
        return 1
    if sid in {"titre2", "heading2"}:
        return 2
    if sid in {"titre3", "heading3"}:
        return 3
    if sid in {"titre4", "heading4"}:
        return 4
    # Unnumbered H1-like
    up = text.upper().strip()
    if up in {
        "RESUME",
        "RÉSUMÉ",
        "ABSTRACT",
        "INTRODUCTION GÉNÉRALE",
        "INTRODUCTION GENERALE",
        "CONCLUSION GENERALE",
        "CONCLUSION GÉNÉRALE",
        "PERSPECTIVES",
        "BIBLIOGRAPHIE",
        "WEBOGRAPHIE",
        "GLOSSAIRE",
        "TABLE DE MATIERES",
        "TABLE DES MATIÈRES",
        "ANNEXES",
    }:
        return 1
    return 0


def _match_content_key(keys: list[str], h1: str, h2: str, h3: str) -> str | None:
    """Find best matching content key for current heading path."""
    candidates = []
    if h3:
        candidates.append(f"{h1}|{h2}|{h3}")
    if h2:
        candidates.append(f"{h1}|{h2}")
    candidates.append(h1)

    def soften(s: str) -> str:
        s = _norm(s).upper()
        s = s.replace("É", "E").replace("È", "E").replace("Ê", "E").replace("À", "A")
        s = s.replace("Ù", "U").replace("Ô", "O").replace("Î", "I").replace("Ç", "C")
        s = re.sub(r"[^A-Z0-9| ]+", "", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    soft_keys = {soften(k): k for k in keys}
    for c in candidates:
        sc = soften(c)
        if sc in soft_keys:
            return soft_keys[sc]
        # partial: endswith
        for sk, orig in soft_keys.items():
            if sk.endswith(sc) or sc.endswith(sk):
                return orig
            # dossier number flexible
            if "|" in sc and "|" in sk:
                parts_c = sc.split("|")
                parts_k = sk.split("|")
                if parts_c[-1] == parts_k[-1] and parts_c[0].split()[:2] == parts_k[0].split()[:2]:
                    return orig
    return None


def _insert_paragraphs_after(anchor: etree._Element, texts: list[str], *, italic: bool = False) -> int:
    parent = anchor.getparent()
    idx = parent.index(anchor)
    n = 0
    for i, text in enumerate(texts):
        new_p = OxmlElement("w:p")
        _set_p_format(new_p)
        new_p.append(_make_run(text, italic=italic or text.startswith("[ INSÉRER")))
        parent.insert(idx + 1 + i, new_p)
        n += 1
    return n


def _set_numpr(p: etree._Element, ilvl: str = "0") -> None:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        numpr = OxmlElement("w:numPr")
        ppr.append(numpr)
    il = numpr.find(qn("w:ilvl"))
    if il is None:
        il = OxmlElement("w:ilvl")
        numpr.append(il)
    il.set(qn("w:val"), ilvl)
    nid = numpr.find(qn("w:numId"))
    if nid is None:
        nid = OxmlElement("w:numId")
        numpr.append(nid)
    nid.set(qn("w:val"), "1")


def _fill_table_cells(tbl: etree._Element, rows_data: list[list[str]]) -> None:
    rows = tbl.findall(qn("w:tr"))
    for ri, row_data in enumerate(rows_data):
        ridx = ri + 1  # skip header
        if ridx >= len(rows):
            break
        cells = rows[ridx].findall(qn("w:tc"))
        for ci, val in enumerate(row_data):
            if ci >= len(cells):
                break
            # rewrite first paragraph of cell
            ps = cells[ci].findall(qn("w:p"))
            if not ps:
                continue
            _rewrite_p(ps[0], val)
            # clear extra paragraphs text
            for extra in ps[1:]:
                _clear_p(extra)


def _heading_level_name(style_id: str, text: str) -> tuple[int, str]:
    lvl = _is_heading(style_id, text)
    return lvl, _norm(text)


def fill_document_xml(doc_root: etree._Element, data: dict[str, Any]) -> dict[str, int]:
    stats = {
        "placeholders_filled": 0,
        "sections_appended": 0,
        "bullets_added": 0,
        "tables_filled": 0,
        "diagram_placeholders": 0,
        "end_matter": 0,
    }
    content: dict[str, list[str]] = data["content"]
    bullets: dict[str, list[str]] = data["bullets"]
    tables = data["tables"]

    # Normalize content keys once
    content_keys = list(content.keys())
    bullet_keys = list(bullets.keys())

    body = doc_root.find(qn("w:body"))
    if body is None:
        return stats

    children = list(body)
    h1 = h2 = h3 = ""
    protected = False
    filled_keys: set[str] = set()
    # Track tables order for cost/critique
    table_index = 0
    pending_table_role: str | None = None

    i = 0
    while i < len(children):
        el = children[i]
        if el.tag == qn("w:tbl"):
            if pending_table_role and pending_table_role in tables:
                _fill_table_cells(el, tables[pending_table_role])
                stats["tables_filled"] += 1
                pending_table_role = None
            table_index += 1
            i += 1
            continue

        if el.tag != qn("w:p"):
            i += 1
            continue

        text = _para_text(el).strip()
        style = _style_id(el)
        lvl, ntext = _heading_level_name(style, text)

        if lvl == 1 and ntext:
            h1, h2, h3 = ntext, "", ""
            protected = any(_norm(p).upper() in _norm(h1).upper() or _norm(h1).upper().startswith(_norm(p).upper()[:20]) for p in PROTECTED_H1)
            # stricter
            up = h1.upper()
            protected = (
                up.startswith("DÉDICACE")
                or up.startswith("DEDICACE")
                or up.startswith("REMERCIEMENTS")
                or up.startswith("PARTIE I")
            )
            # End-matter empty headings: append content
            key = _match_content_key(content_keys, h1, "", "")
            if key and key not in filled_keys and not protected:
                # If next non-empty body para is missing (heading ≠ body), append
                nxt_text = ""
                nxt_is_heading = False
                for j in range(i + 1, min(i + 8, len(children))):
                    if children[j].tag != qn("w:p"):
                        continue
                    cand = _para_text(children[j]).strip()
                    st = _style_id(children[j])
                    if not cand:
                        continue
                    if _is_heading(st, cand):
                        nxt_is_heading = True
                        nxt_text = cand
                    else:
                        nxt_text = cand
                        nxt_is_heading = False
                    break
                needs = (
                    not nxt_text
                    or nxt_is_heading
                    or PLACEHOLDER_RE.match(nxt_text)
                    or nxt_text.startswith("[À COMPLÉTER")
                )
                # Ne pas re-remplir RESUME déjà rédigé
                already = bool(nxt_text) and not nxt_is_heading and (
                    "MOOMOO" in nxt_text
                    or "World Health" in nxt_text
                    or "internship report" in nxt_text.lower()
                )
                if needs and not already and key in content:
                    n = _insert_paragraphs_after(el, content[key])
                    stats["sections_appended"] += n
                    stats["end_matter"] += 1
                    filled_keys.add(key)
                    children = list(body)
                    i += 1 + n
                    continue
                if already:
                    filled_keys.add(key)

        elif lvl == 2 and ntext:
            h2, h3 = ntext, ""
        elif lvl == 3 and ntext:
            h3 = ntext

        # Detect table roles from headings
        up2 = (h2 + " " + h3).upper()
        if "CRITIQUE" in up2 and "EXISTANT" in (h1 + h2).upper():
            pending_table_role = "critique"
        elif "RESSOURCES MATERIELLES" in up2.replace("É", "E") or "RESSOURCES MATÉRIELLES" in (h2 + h3).upper():
            pending_table_role = "materiel"
        elif "RESSOURCES LOGICIELLES" in up2:
            pending_table_role = "logiciel"
        elif "RESSOURCES HUMAINES" in up2:
            pending_table_role = "rh"
        elif "COUT TOTAL" in up2.replace("Û", "U") or "COÛT TOTAL" in (h2 + h3).upper() or ntext.upper().startswith("COUT TOTAL"):
            pending_table_role = "total"

        if protected:
            i += 1
            continue

        # Fill placeholder paragraphs
        if text and PLACEHOLDER_RE.match(text):
            key = _match_content_key(content_keys, h1, h2, h3)
            if key and key in content:
                paras = content[key]
                _rewrite_p(el, paras[0], italic=paras[0].startswith("[ INSÉRER"))
                if len(paras) > 1:
                    n = _insert_paragraphs_after(el, paras[1:])
                    stats["sections_appended"] += n
                    children = list(body)
                    i += 1 + n
                else:
                    i += 1
                stats["placeholders_filled"] += 1
                filled_keys.add(key)
                # Bullets after section?
                bkey = _match_content_key(bullet_keys, h1, h2, h3)
                if bkey and bkey in bullets:
                    # insert after the block we just added
                    # find last inserted = current position after inserts
                    anchor = children[min(i, len(children) - 1)] if i < len(children) else el
                    # better: insert after el's last following content para
                    # Use el as anchor if single para else the last inserted
                    # Re-get el index
                    # Insert bullets after filled block
                    # Find el again
                    pass
                continue
            # Diagram-ish placeholder
            if DIAGRAM_HINTS.search(h2 + h3 + text) or "diagramme" in text.lower() or "gant" in text.lower():
                name = h3 or h2 or "Diagramme MOOMOO"
                _rewrite_p(el, f"[ INSÉRER DIAGRAMME ICI : {name} ]", italic=True)
                stats["diagram_placeholders"] += 1
                i += 1
                continue
            # Generic leftover
            _rewrite_p(
                el,
                f"[À COMPLÉTER — section « {_norm(h2 or h1)} » : contenu métier MOOMOO non fourni]",
            )
            stats["placeholders_filled"] += 1
            i += 1
            continue

        i += 1

    # Second pass: add bullet lists where section filled but bullets missing
    stats["bullets_added"] += _ensure_bullets(body, bullets)
    # Aggressive sweep: remaining Insérer… hors PARTIE I / dédicace / remerciements
    swept = _sweep_remaining_placeholders(body, content)
    stats["placeholders_filled"] += swept
    # Uniformize any remaining "Insérez votre diagramme"
    stats["diagram_placeholders"] += _uniform_diagram_placeholders(body)
    # End matter: biblio, web, glossaire, annexes
    stats["end_matter"] += _fill_end_matter(body, data)
    # Source lines under cost tables
    stats["sections_appended"] += _add_cost_sources(body)
    # Déduplication de paragraphes corps identiques consécutifs (ex. Gantt doublé)
    stats["deduped"] = _dedupe_consecutive_paras(body)
    stats["stack_caveats"] = _mark_stack_hypotheses(body)

    return stats


def _mark_stack_hypotheses(body: etree._Element) -> int:
    """Signale React/FastAPI/Gemini comme hypothèses non confirmées pour MOOMOO."""
    n = 0
    needles = ("React", "FastAPI", "Gemini")
    for p in body.iter(qn("w:p")):
        text = _para_text(p)
        if not any(x in text for x in needles):
            continue
        if "[À VÉRIFIER : stack gabarit" in text or "[À VÉRIFIER : hypothèse gabarit" in text:
            continue
        # Uniquement les mentions d'implémentation (pas la biblio web)
        if not any(
            k in text
            for k in (
                "développ",
                "implément",
                "Couche",
                "Backend",
                "Frontend",
                "intégration",
            )
        ):
            continue
        new_text = text
        if "React" in new_text and "stack gabarit" not in new_text:
            new_text = new_text.replace(
                "React", "React [À VÉRIFIER : stack gabarit, non confirmée pour MOOMOO]", 1
            )
        if "FastAPI" in new_text and "stack gabarit" not in new_text:
            new_text = new_text.replace(
                "FastAPI",
                "FastAPI [À VÉRIFIER : stack gabarit, non confirmée pour MOOMOO]",
                1,
            )
        if "Gemini" in new_text and "hypothèse gabarit" not in new_text:
            new_text = new_text.replace(
                "Gemini", "Gemini [À VÉRIFIER : hypothèse gabarit]", 1
            )
        if new_text != text:
            _rewrite_p(p, new_text)
            n += 1
    return n


def _dedupe_consecutive_paras(body: etree._Element) -> int:
    n = 0
    # 1) doublons stricts consécutifs
    children = list(body)
    prev_text = None
    for el in children:
        if el.tag != qn("w:p"):
            prev_text = None
            continue
        text = _para_text(el).strip()
        style = _style_id(el)
        if _is_heading(style, text):
            prev_text = None
            continue
        if text and prev_text and text == prev_text:
            parent = el.getparent()
            parent.remove(el)
            n += 1
            continue
        if text:
            prev_text = text
    # 2) blocs répétés en ignorant les paragraphes vides (ex. Gantt doublé)
    def _nonempty_paras() -> list[etree._Element]:
        out: list[etree._Element] = []
        for c in list(body):
            if c.tag != qn("w:p"):
                continue
            t = _para_text(c).strip()
            st = _style_id(c)
            if not t or _is_heading(st, t):
                continue
            out.append(c)
        return out

    changed = True
    while changed:
        changed = False
        ne = _nonempty_paras()
        i = 0
        while i + 6 <= len(ne):
            a = [_para_text(ne[i + k]).strip() for k in range(3)]
            b = [_para_text(ne[i + 3 + k]).strip() for k in range(3)]
            if a == b:
                parent = ne[i].getparent()
                for k in range(2, -1, -1):
                    parent.remove(ne[i + 3 + k])
                    n += 1
                changed = True
                break
            i += 1
    return n


def _sweep_remaining_placeholders(body: etree._Element, content: dict[str, list[str]]) -> int:
    """Remplace tout Insérer/Insérez restant hors zones protégées."""
    n = 0
    h1 = h2 = h3 = ""
    protected = False
    content_keys = list(content.keys())
    children = list(body)
    i = 0
    while i < len(children):
        el = children[i]
        if el.tag != qn("w:p"):
            i += 1
            continue
        text = _para_text(el).strip()
        style = _style_id(el)
        lvl, ntext = _heading_level_name(style, text)
        if lvl == 1 and ntext:
            h1, h2, h3 = ntext, "", ""
            up = h1.upper()
            protected = (
                up.startswith("DÉDICACE")
                or up.startswith("DEDICACE")
                or up.startswith("REMERCIEMENTS")
                or up.startswith("PARTIE I")
            )
        elif lvl == 2 and ntext:
            h2, h3 = ntext, ""
        elif lvl == 3 and ntext:
            h3 = ntext

        if protected:
            i += 1
            continue

        is_ph = bool(text and PLACEHOLDER_RE.match(text))
        if not is_ph and text and re.match(r"^Ins[eé]r", text, re.I):
            is_ph = True
        if not is_ph:
            i += 1
            continue

        key = _match_content_key(content_keys, h1, h2, h3)
        if key and key in content:
            paras = content[key]
            nxt = ""
            if i + 1 < len(children) and children[i + 1].tag == qn("w:p"):
                nxt = _para_text(children[i + 1]).strip()
            if nxt and paras and nxt[:60] == paras[0][:60]:
                # Placeholder redondant : le contenu est déjà juste après
                parent = el.getparent()
                parent.remove(el)
                children = list(body)
                n += 1
                # ne pas avancer i : la liste a été reconstruite, reprendre au même index
                continue
            # Si un des paragraphes suivants (même section) contient déjà le début du contenu
            already_block = False
            for k in range(i + 1, min(i + 6, len(children))):
                if children[k].tag != qn("w:p"):
                    continue
                tk = _para_text(children[k]).strip()
                st = _style_id(children[k])
                if tk and _is_heading(st, tk):
                    break
                if paras and tk[:50] == paras[0][:50]:
                    already_block = True
                    break
            if already_block:
                parent = el.getparent()
                parent.remove(el)
                children = list(body)
                n += 1
                continue
            _rewrite_p(el, paras[0], italic=paras[0].startswith("[ INSÉRER"))
            extra = paras[1:]
            if nxt and extra and nxt[:50] == extra[0][:50]:
                extra = []
            if extra:
                added = _insert_paragraphs_after(el, extra)
                children = list(body)
                i += 1 + added
            else:
                i += 1
            n += 1
            continue

        if DIAGRAM_HINTS.search(h2 + " " + h3 + " " + text) or re.search(
            r"diagramme|figure|organigramme|capture|gantt", text, re.I
        ):
            name = h3 or h2 or "Diagramme MOOMOO"
            _rewrite_p(el, f"[ INSÉRER DIAGRAMME ICI : {name} ]", italic=True)
            n += 1
            i += 1
            continue

        _rewrite_p(
            el,
            f"[À COMPLÉTER — section « {_norm(h2 or h3 or h1)} » : contenu métier MOOMOO non fourni]",
        )
        n += 1
        i += 1
    return n


def _ensure_bullets(body: etree._Element, bullets: dict[str, list[str]]) -> int:
    """Injecte les listes à puces une seule fois par clé, sous le bon titre."""
    n = 0
    h1 = h2 = h3 = ""
    jobs: list[tuple[etree._Element, str]] = []
    seen_keys: set[str] = set()
    for el in list(body):
        if el.tag != qn("w:p"):
            continue
        text = _para_text(el).strip()
        style = _style_id(el)
        lvl, ntext = _heading_level_name(style, text)
        if lvl == 1 and ntext:
            h1, h2, h3 = ntext, "", ""
        elif lvl == 2 and ntext:
            h2, h3 = ntext, ""
        elif lvl == 3 and ntext:
            h3 = ntext
        else:
            continue
        if lvl not in (2, 3):
            continue
        key = _match_content_key(list(bullets.keys()), h1, h2, h3)
        if not key or key in seen_keys:
            continue
        # Vérifier que la clé correspond bien au titre courant (évite faux positifs)
        key_tail = key.split("|")[-1].upper()
        if _norm(ntext).upper().replace("É", "E") != key_tail.replace("É", "E"):
            # autoriser match si le titre est contenu dans la queue
            if key_tail not in _norm(ntext).upper().replace("É", "E"):
                continue
        seen_keys.add(key)
        jobs.append((el, key))

    for heading_el, key in jobs:
        first_bullet = bullets[key][0][:40]
        # Parcourir les frères suivants jusqu'au prochain titre
        already = False
        insert_at = heading_el
        sib = heading_el.getnext()
        steps = 0
        while sib is not None and steps < 30:
            steps += 1
            if sib.tag != qn("w:p"):
                sib = sib.getnext()
                continue
            t = _para_text(sib).strip()
            st = _style_id(sib)
            if t and _is_heading(st, t):
                break
            if t and (first_bullet in t or t.startswith("BF01") or t.startswith("BNF01") or "Utilisateur sourd" in t):
                already = True
                break
            if t:
                insert_at = sib
            sib = sib.getnext()
        if already:
            continue
        for bi, btext in enumerate(bullets[key]):
            new_p = OxmlElement("w:p")
            _set_p_format(new_p)
            _set_numpr(new_p, "0")
            new_p.append(_make_run(btext))
            parent = insert_at.getparent()
            parent.insert(parent.index(insert_at) + 1 + bi, new_p)
            n += 1
    return n


def _uniform_diagram_placeholders(body: etree._Element) -> int:
    n = 0
    for p in body.iter(qn("w:p")):
        text = _para_text(p).strip()
        if re.search(r"Ins[eé]rez votre diagramme", text, re.I):
            _rewrite_p(p, "[ INSÉRER DIAGRAMME ICI : Diagramme de Gantt MOOMOO ]", italic=True)
            n += 1
        elif re.search(r"\[À COMPLÉTER : cas d'utilisation MOOMOO\]", text) and "DIAGRAMME" in text.upper():
            # Add companion placeholder para if not present — rewrite to include insert marker
            if "[ INSÉRER DIAGRAMME ICI" not in text:
                _rewrite_p(
                    p,
                    text + " — [ INSÉRER DIAGRAMME ICI : Cas d'utilisation MOOMOO ]",
                    italic=False,
                )
                n += 1
    # After headings that clearly need diagrams (Cas 1/2/3 empty)
    h2 = h3 = h1 = ""
    children = list(body)
    for idx, el in enumerate(children):
        if el.tag != qn("w:p"):
            continue
        text = _para_text(el).strip()
        style = _style_id(el)
        lvl, ntext = _heading_level_name(style, text)
        if lvl == 1 and ntext:
            h1, h2, h3 = ntext, "", ""
        elif lvl == 2 and ntext:
            h2, h3 = ntext, ""
        elif lvl == 3 and ntext:
            h3 = ntext
        elif lvl == 4 and ntext:
            # Cas 1/2/3 under sequence/activity
            if re.match(r"^Cas\s*[123]$", ntext, re.I) and any(
                x in (h1 + h2 + h3).upper() for x in ("SEQUENCE", "SÉQUENCE", "ACTIVITE", "ACTIVITÉ", "COMMUNICATION")
            ):
                # if next is empty/placeholder
                empty = True
                for j in range(idx + 1, min(idx + 3, len(children))):
                    if children[j].tag != qn("w:p"):
                        continue
                    t = _para_text(children[j]).strip()
                    if t and not PLACEHOLDER_RE.match(t) and "[ INSÉRER" not in t and "[À COMPLÉTER" not in t:
                        empty = False
                    break
                if empty:
                    label = f"{h3 or h2} — {ntext}"
                    # check if already inserted
                    nxt = _para_text(children[idx + 1]).strip() if idx + 1 < len(children) else ""
                    if "[ INSÉRER DIAGRAMME ICI" not in nxt:
                        _insert_paragraphs_after(
                            el,
                            [f"[ INSÉRER DIAGRAMME ICI : {label} ]"],
                        )
                        n += 1
                        children = list(body)
    return n


def _fill_end_matter(body: etree._Element, data: dict[str, Any]) -> int:
    n = 0
    biblio = json.loads(REFS_BIBLIO.read_text(encoding="utf-8"))
    web = json.loads(REFS_WEB.read_text(encoding="utf-8"))
    glossaire = data["glossaire"]

    targets = {
        "BIBLIOGRAPHIE": [f"{i+1}. {b['apa']}" for i, b in enumerate(biblio)],
        "WEBOGRAPHIE": [
            f"{i+1}. {w['title']}. {w['url']} (consulté le {w['accessed']})."
            for i, w in enumerate(web)
        ],
        "GLOSSAIRE": [f"{k} : {v}" for k, v in glossaire],
    }

    children = list(body)
    for idx, el in enumerate(children):
        if el.tag != qn("w:p"):
            continue
        text = _norm(_para_text(el)).upper()
        for key, lines in targets.items():
            if text == key or text.startswith(key):
                # Skip if already filled
                nxt = ""
                if idx + 1 < len(children) and children[idx + 1].tag == qn("w:p"):
                    nxt = _para_text(children[idx + 1]).strip()
                if nxt and (
                    "World Health Organization" in nxt
                    or "ASL :" in nxt
                    or nxt.strip()[:4].rstrip(".").isdigit()
                    or "http" in nxt
                ):
                    continue
                n += _insert_paragraphs_after(el, lines)
                children = list(body)
                break

    # ANNEXES — insert before BIBLIOGRAPHIE if missing
    has_annexes = any(
        _norm(_para_text(p)).upper().startswith("ANNEXES")
        for p in body.iter(qn("w:p"))
        if _is_heading(_style_id(p), _para_text(p).strip())
    )
    if not has_annexes:
        for p in list(body):
            if p.tag != qn("w:p"):
                continue
            if _norm(_para_text(p)).upper() == "BIBLIOGRAPHIE":
                # Insert ANNEXES heading + content before biblio
                h = OxmlElement("w:p")
                ppr = OxmlElement("w:pPr")
                ps = OxmlElement("w:pStyle")
                ps.set(qn("w:val"), "Titre1")
                ppr.append(ps)
                h.append(ppr)
                _set_p_format(h, center=True)
                h.append(_make_run("ANNEXES"))
                parent = p.getparent()
                parent.insert(parent.index(p), h)
                annex_lines = [
                    "Annexe A — Crédits et sources des images / diagrammes",
                    "Les figures et diagrammes encore absents sont signalés par le marqueur uniforme "
                    "« [ INSÉRER DIAGRAMME ICI : Nom ] ». Toute image tierce éventuelle devra préciser : "
                    "auteur, licence, URL et date de consultation.",
                    "Annexe B — Documents administratifs de stage",
                    "Lettre d'admission en stage signée, fiches de stage et annexes du cahier des charges IAI "
                    ": [À COMPLÉTER — pièces à joindre].",
                    "Annexe C — Captures d'écran de l'application MOOMOO",
                    "[ INSÉRER DIAGRAMME ICI : Jeu de captures d'écran MOOMOO — annexe C ]",
                    "Note : le cahier des charges IAI place bibliographie/webographie tantôt dans les annexes, "
                    "tantôt en sections distinctes ; le présent rapport conserve des sections séparées "
                    "BIBLIOGRAPHIE / WEBOGRAPHIE / GLOSSAIRE et regroupe ici les annexes documentaires "
                    "(ambiguïté signalée dans A_VERIFIER.md).",
                ]
                n += _insert_paragraphs_after(h, annex_lines)
                n += 1
                break
    return n


def _add_cost_sources(body: etree._Element) -> int:
    n = 0
    data = _load_data()
    src = data["sources_inline"]
    markers = [
        (
            "RESSOURCES MATERIELLES",
            f"Source : prix catalogue Zoom Informatique Cameroun (HP 240 G10 = 250 000 FCFA ; "
            f"Lenovo IdeaPad 3 = 210 000 FCFA), consulté le 29/09/2026. "
            f"Référentiel : Mercuriale des prix 2026 (MINCOMMERCE/BOND) — {src['merc']}. "
            f"Autres lignes : [À VÉRIFIER].",
        ),
        (
            "RESSOURCES LOGICIELLES",
            "Source : logiciels open source à coût de licence 0 (Python.org, MediaPipe/Google AI Edge, "
            "PostgreSQL.org — consultation 29/09/2026). Coûts d'hébergement cloud éventuels : [À VÉRIFIER].",
        ),
        (
            "RESSOURCES HUMAINES",
            f"Source fourchettes salaires développeur junior Cameroun 2026 : 120 000–200 000 FCFA/mois "
            f"(ECONUMA, 2026) et 150 000–300 000 FCFA/mois (Sango Services, 2026). "
            f"Indemnité de stage et durée d'annotation : [À VÉRIFIER] — aucun total inventé.",
        ),
        (
            "COUT TOTAL",
            "Source : consolidation des lignes précédentes. Les totaux non entièrement sourcés restent "
            "[À VÉRIFIER] volontairement (règle : ne pas inventer de prix).",
        ),
    ]
    children = list(body)
    h_stack = ""
    for idx, el in enumerate(children):
        if el.tag != qn("w:p"):
            continue
        text = _para_text(el).strip()
        style = _style_id(el)
        lvl, ntext = _heading_level_name(style, text)
        if lvl and ntext:
            soft = ntext.upper().replace("É", "E").replace("È", "E").replace("Û", "U")
            for key, source_line in markers:
                if key in soft:
                    # Find end of following table
                    j = idx + 1
                    last = el
                    while j < len(children):
                        if children[j].tag == qn("w:tbl"):
                            last = children[j]
                            j += 1
                            # after table, insert source if not present
                            nxt = ""
                            if j < len(children) and children[j].tag == qn("w:p"):
                                nxt = _para_text(children[j])
                            if not nxt.lower().startswith("source"):
                                new_p = OxmlElement("w:p")
                                _set_p_format(new_p)
                                new_p.append(_make_run(source_line, italic=True, size_pt=FONT_CAPTION_PT))
                                parent = last.getparent()
                                parent.insert(parent.index(last) + 1, new_p)
                                n += 1
                                children = list(body)
                            break
                        if children[j].tag == qn("w:p"):
                            t2 = _para_text(children[j]).strip()
                            st2 = _style_id(children[j])
                            if _is_heading(st2, t2):
                                break
                        j += 1
    return n


def write_side_docs_passes(stats: dict) -> None:
    biblio = json.loads(REFS_BIBLIO.read_text(encoding="utf-8"))
    web = json.loads(REFS_WEB.read_text(encoding="utf-8"))

    a_lines = [
        "# A_VERIFIER — points à compléter (PASSES 2–4)",
        "",
        "## Bloc utilisateur incomplet (ne pas inventer)",
        "- Superviseur IAI : `[À COMPLÉTER]`",
        "- Maître de stage : `[À COMPLÉTER]`",
        "- Période de stage : `[À COMPLÉTER]`",
        "- Nature / description juridique MOOMOO : `[À COMPLÉTER]`",
        "- Langue des signes ciblée : `[À COMPLÉTER]`",
        "- Plateforme (mobile/web) : `[À COMPLÉTER]`",
        "- Stack technique retenue : `[À VÉRIFIER]` (mentions React/FastAPI/PostgreSQL/Gemini = hypothèses de gabarit)",
        "- Avancement / dataset : `[À COMPLÉTER]`",
        "- Indemnité de stage et coûts RH exacts : `[À VÉRIFIER]`",
        "- Webcam / Internet / smartphone : prix unitaires `[À VÉRIFIER]` (mercuriale 2026 consultable)",
        "",
        "## Ambiguïté cahier des charges (annexes / conclusion)",
        "- Le CDC IAI (scan CamScanner) liste bibliographie/webographie/glossaire tantôt **dans les annexes**,",
        "  tantôt comme sections de fin de rapport ; la page « plan du rapport » place aussi la table des matières en fin d'annexes.",
        "- **Choix retenu** : sections séparées CONCLUSION → PERSPECTIVES → ANNEXES → BIBLIOGRAPHIE → WEBOGRAPHIE → GLOSSAIRE → TABLE DE MATIERES,",
        "  avec note explicite dans ANNEXES. À faire valider par le superviseur.",
        "",
        "## Prix 2026 FCFA sourcés (retenus)",
        "- PC HP 240 G10 i3 : **250 000 FCFA** — https://zoominformatics.com/produit/laptop-hp-240-g10-intel-core-i3-1315u/",
        "- PC Lenovo IdeaPad 3 : **210 000 FCFA** — https://zoominformatics.com/produit/pc-portable-lenovo-ideapad-3/",
        "- Mercuriale 2026 (référentiel) : https://lefisk.cm/outils/mercuriale",
        "- Salaire junior (fourchettes) : ECONUMA 120–200 k ; Sango Services 150–300 k FCFA/mois — pas de total RH inventé",
        "",
        "## Stats passes 2–4",
        f"```",
        f"{json.dumps(stats, ensure_ascii=False, indent=2)}",
        f"```",
        "",
        f"## Références chargées : {len(biblio)} biblio, {len(web)} web",
        "",
    ]
    A_VERIFIER.write_text("\n".join(a_lines), encoding="utf-8")

    # Collect diagram placeholders by scanning OUT if exists
    p_lines = [
        "# PLACEHOLDERS — inventaire (PASSES 2–4)",
        "",
        "## Conservés (ne pas remplir — consignes)",
        "- PARTIE I : `Insérer votre texte ici…`, `NOM DE LA STRUTURE`, etc.",
        "- DÉDICACE / REMERCIEMENTS : contenu intact",
        "",
        "## Marqueurs métier",
        "- `[À COMPLÉTER]` — fait propre au stage / entreprise / dataset non fourni",
        "- `[À VÉRIFIER]` — hypothèse ou prix/stack non confirmé",
        "- `[ INSÉRER DIAGRAMME ICI : Nom ]` — figure/capture/diagramme manquant (format uniforme)",
        "- Champs TOC/SEQ : rafraîchir dans Word / LibreOffice (`F9` / mise à jour des champs)",
        "",
        "## Diagrammes / captures attendus (liste type)",
        "- Diagramme de Gantt MOOMOO",
        "- Cas d'utilisation, communication, séquence, activité (Cas 1–3)",
        "- Diagramme de classes / états / paquetages",
        "- Déploiement / composants",
        "- Captures installation et écrans utilisateur",
        "- Annexe C — jeu de captures",
        "",
        f"Stats: {stats}",
        "",
    ]
    PLACEHOLDERS.write_text("\n".join(p_lines), encoding="utf-8")


def apply_passes(doc_path: Path) -> dict:
    data = _load_data()
    # Work on raw XML for robustness
    import zipfile
    from io import BytesIO

    with zipfile.ZipFile(doc_path, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    doc_root = etree.fromstring(parts["word/document.xml"])
    stats = fill_document_xml(doc_root, data)
    parts["word/document.xml"] = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, raw in parts.items():
            zout.writestr(name, raw)
    doc_path.write_bytes(buf.getvalue())

    # python-docx pass: fonts on new paras
    doc = Document(str(doc_path))
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = FONT_NAME
            rpr = run._element.get_or_add_rPr()
            _ensure_rfonts(rpr)
    doc.save(str(doc_path))

    write_side_docs_passes(stats)
    return stats


if __name__ == "__main__":
    from pathlib import Path as P

    apply_passes(P("/workspace/Mon_rapport_COMPLET.docx"))
