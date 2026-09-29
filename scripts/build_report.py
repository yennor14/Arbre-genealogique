#!/usr/bin/env python3
"""
PASSES 1–4 — Normalisation + rédaction du rapport de stage MOOMOO.

Idempotent : repart toujours de SRC et réécrit OUT, puis applique passes 2–4.
Usage : python3 scripts/build_report.py
"""
from __future__ import annotations

import re
import shutil
import zipfile
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree

# ---------------------------------------------------------------------------
# Constantes de mise en forme (PASSE 1)
# ---------------------------------------------------------------------------
SRC = Path("/workspace/Mon_rapport.docx")
OUT = Path("/workspace/Mon_rapport_COMPLET.docx")
RAPPORT_MD = Path("/workspace/rapport.md")
A_VERIFIER = Path("/workspace/A_VERIFIER.md")
PLACEHOLDERS = Path("/workspace/PLACEHOLDERS.md")

TITLE_EXACT = (
    "CONCEPTION ET RÉALISATION D'UNE APPLICATION DE TRADUCTION "
    "DU LANGAGE DES SIGNES EN TEXTE OU AUDIO ET VICE VERSA : CAS DE MOOMOO"
)
# En-tête : titre exact (lisible, sans troncature brutale)
HEADER_TITLE = TITLE_EXACT

FONT_NAME = "Times New Roman"
FONT_BODY_PT = 12
FONT_TABLE_PT = 11
FONT_CAPTION_PT = 10
LINE_SPACING = 1.5  # -> w:line=360, lineRule=auto
LINE_TWIPS = 360
TABLE_LINE_SPACING = 1.0
TABLE_LINE_TWIPS = 240
SPACING_AFTER_PT = 6
MARGIN_CM = 2.5
PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
USEFUL_WIDTH_CM = 16.0  # 21 - 2*2.5

COLOR_PALETTE = {
    "2F5496",
    "1F3763",
    "FF6600",
    "FF9966",
    "44546A",
    "7030A0",
    "E4CCFC",
    "4472C4",
    "0563C1",
    "FFFFFF",
}
TABLE_HEADER_FILL = "7030A0"
TABLE_ALT_FILL = "E4CCFC"
HEADING1_COLOR = "2F5496"
HEADING3_COLOR = "1F3763"

COMPAT_MODE = "15"
LIST_NUM_ID = "1"
LIST_ABSTRACT_ID = "0"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NSMAP = {"w": W_NS, "r": R_NS}

# Remplacements texte (ordre important). Ne pas toucher placeholders PARTIE I.
TYPO_REPLACEMENTS: list[tuple[str, str]] = [
    ("NARY JACOBSON", "IVAR JACOBSON"),
    ("Nary Jacobson", "Ivar Jacobson"),
    ("JAMES RUMBAUCH", "JAMES RUMBAUGH"),
    ("James Rumbauch", "James Rumbaugh"),
    ("RÉSENTATION DES ARCHITECTURES", "PRÉSENTATION DES ARCHITECTURES"),
    ("DECRIPTION TEXTUELLE", "DESCRIPTION TEXTUELLE"),
    ("Digramme de communication", "Diagramme de communication"),
    ("DIGRAMME DE COMMUNICATION", "DIAGRAMME DE COMMUNICATION"),
    ("VIS VERSA", "VICE VERSA"),
    ("REALISATION", "RÉALISATION"),
]

# Remplacements regex (évite Gant → Ganttt sur un Gantt déjà correct)
TYPO_REGEX_REPLACEMENTS: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"(?i)(diagramme de )Gant(?!t)"), r"\1Gantt"),
]

RESIDUE_REPLACEMENTS: list[tuple[str, str]] = [
    ("AZOUTSA", "[À COMPLÉTER]"),
    ("06/08/2026", "[À COMPLÉTER]"),
    (
        "application web d’orientation intelligente et de mentorat",
        "application [À COMPLÉTER : MOOMOO]",
    ),
    (
        "application d’orientation intelligente et de mentorat",
        "application [À COMPLÉTER : MOOMOO]",
    ),
    (
        "plateforme d’orientation intelligente et de mentorat",
        "plateforme [À COMPLÉTER : MOOMOO]",
    ),
    (
        "étudiants, mentors, professionnels, administrateurs et validateurs",
        "[À COMPLÉTER : acteurs MOOMOO]",
    ),
    ("pour Nexura", "pour [À COMPLÉTER : MOOMOO]"),
    ("Nexura", "[À COMPLÉTER]"),
]

# Titres / légendes clairement issus de l’ancien projet mentorat
RESIDUE_LINE_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (
        re.compile(
            r"DIAGRAMME DE CAS D’UTILISATION\s*«\s*Gérer ses séances\s*»",
            re.I,
        ),
        "DIAGRAMME DE CAS D’UTILISATION « [À COMPLÉTER : cas d’utilisation MOOMOO] »",
    ),
    (
        re.compile(
            r'Figure\s*7\s*:\s*Diagramme de cas d\'utilisation\s*"\s*interagir à une publication\s*"',
            re.I,
        ),
        "Figure 7: Diagramme de cas d'utilisation « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r"DIAGRAMME DE CAS D’UTILISATION\s*«\s*GERER SES PUBLICATIONS\s*»",
            re.I,
        ),
        "DIAGRAMME DE CAS D’UTILISATION « [À COMPLÉTER : cas d’utilisation MOOMOO] »",
    ),
    (
        re.compile(
            r'Figure\s*8\s*:\s*Diagramme de cas d\'utilisation\s*"\s*Gérer ses publications\s*"',
            re.I,
        ),
        "Figure 8: Diagramme de cas d'utilisation « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(r"DECRIPTION TEXTUELLE\s*:\s*Faire une publication", re.I),
        "DESCRIPTION TEXTUELLE : [À COMPLÉTER — cas d’utilisation MOOMOO]",
    ),
    (
        re.compile(r"DESCRIPTION TEXTUELLE\s*:\s*Faire une publication", re.I),
        "DESCRIPTION TEXTUELLE : [À COMPLÉTER — cas d’utilisation MOOMOO]",
    ),
    (
        re.compile(r"DESCRIPTION TEXTUELLE\s*:\s*Commenter une publication", re.I),
        "DESCRIPTION TEXTUELLE : [À COMPLÉTER — cas d’utilisation MOOMOO]",
    ),
    (
        re.compile(
            r"Tableau\s*16\s*:\s*Description textuelle Faire une publication",
            re.I,
        ),
        "Tableau 16: Description textuelle [À COMPLÉTER — cas d’utilisation MOOMOO]",
    ),
    (
        re.compile(
            r"Tableau\s*17\s*:\s*Description textuelle Commenter une publication",
            re.I,
        ),
        "Tableau 17: Description textuelle [À COMPLÉTER — cas d’utilisation MOOMOO]",
    ),
    (
        re.compile(
            r"Digramme de communication de\s*«\s*faire une publication\s*»",
            re.I,
        ),
        "Diagramme de communication de « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r"Diagramme de communication de\s*«\s*faire une publication\s*»",
            re.I,
        ),
        "Diagramme de communication de « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r"Digramme de communication de\s*«\s*commenter une publication\s*»",
            re.I,
        ),
        "Diagramme de communication de « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r"Diagramme de communication de\s*«\s*commenter une publication\s*»",
            re.I,
        ),
        "Diagramme de communication de « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r'Figure\s*11\s*:\s*Diagramme de communication\s*"\s*Faire une publication\s*"',
            re.I,
        ),
        "Figure 11: Diagramme de communication « [À COMPLÉTER : MOOMOO] »",
    ),
    (
        re.compile(
            r'Figure\s*12\s*:\s*Diagramme de communication\s*"\s*commenter une publication\s*"',
            re.I,
        ),
        "Figure 12: Diagramme de communication « [À COMPLÉTER : MOOMOO] »",
    ),
]

CLASS_ASSOC_RE = re.compile(
    r"\b(Apprenant|Mentor|Validateur|DemandeComptePro|Publication|Commentaire|"
    r"Visioconference|Disponibilites|ConversationMentor|Concours|Retrait|"
    r"Abonnement|Convention)\b",
    re.I,
)
MENTORAT_PARA_RE = re.compile(
    r"orientation intelligente|mentorat|\bMentor\b|\bApprenant\b|\bValidateur\b|"
    r"\bProfessionnel",
    re.I,
)

# ---------------------------------------------------------------------------
# Helpers OOXML
# ---------------------------------------------------------------------------


def _ensure_rfonts(rpr: etree._Element, font: str = FONT_NAME) -> None:
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{attr}"), font)
    # Retirer les thèmes qui écrasent Times
    for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
        if qn(f"w:{attr}") in rfonts.attrib:
            del rfonts.attrib[qn(f"w:{attr}")]


def _set_sz(rpr: etree._Element, pt: float) -> None:
    half = str(int(pt * 2))
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set(qn("w:val"), half)


def _get_or_create_rpr(parent: etree._Element) -> etree._Element:
    rpr = parent.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        parent.insert(0, rpr)
    return rpr


def _get_or_create_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    return ppr


def _set_spacing(
    ppr: etree._Element,
    *,
    line: int = LINE_TWIPS,
    after_pt: float = SPACING_AFTER_PT,
) -> None:
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:after"), str(int(after_pt * 20)))


def _set_jc(ppr: etree._Element, val: str) -> None:
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), val)


def _make_run(text: str, *, italic: bool = False, size_pt: float = FONT_BODY_PT) -> etree._Element:
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    _ensure_rfonts(rpr)
    _set_sz(rpr, size_pt)
    if italic:
        rpr.append(OxmlElement("w:i"))
        rpr.append(OxmlElement("w:iCs"))
    r.append(rpr)
    t = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def _make_field_runs(instr: str, placeholder: str = "") -> list[etree._Element]:
    """Champ Word simple begin/instr/separate/result/end."""
    runs: list[etree._Element] = []

    def _fld(char_type: str) -> etree._Element:
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        _ensure_rfonts(rpr)
        r.append(rpr)
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), char_type)
        r.append(fc)
        return r

    runs.append(_fld("begin"))
    r_instr = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    _ensure_rfonts(rpr)
    r_instr.append(rpr)
    it = OxmlElement("w:instrText")
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" {instr} "
    r_instr.append(it)
    runs.append(r_instr)
    runs.append(_fld("separate"))
    if placeholder:
        runs.append(_make_run(placeholder, size_pt=FONT_CAPTION_PT, italic=True))
    runs.append(_fld("end"))
    return runs


def _clear_paragraph_content(p: etree._Element) -> etree._Element:
    """Garde pPr, supprime le reste."""
    ppr = p.find(qn("w:pPr"))
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    if ppr is None:
        ppr = _get_or_create_ppr(p)
    return ppr


def _paragraph_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _replace_text_in_element(el: etree._Element, old: str, new: str) -> int:
    n = 0
    for t in el.iter(qn("w:t")):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            n += 1
    return n


def _set_all_rfonts(root: etree._Element, font: str = FONT_NAME) -> None:
    for rfonts in root.iter(qn("w:rFonts")):
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            rfonts.set(qn(f"w:{attr}"), font)
        for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
            key = qn(f"w:{attr}")
            if key in rfonts.attrib:
                del rfonts.attrib[key]
    # Assurer rPr sur chaque run
    for r in root.iter(qn("w:r")):
        rpr = _get_or_create_rpr(r)
        _ensure_rfonts(rpr, font)


# ---------------------------------------------------------------------------
# Étapes
# ---------------------------------------------------------------------------


def export_markdown() -> None:
    import subprocess

    subprocess.run(
        ["pandoc", "-t", "markdown", str(SRC), "-o", str(RAPPORT_MD)],
        check=True,
    )
    print(f"[ok] {RAPPORT_MD}")


def apply_text_cleanups(doc_root: etree._Element) -> dict[str, int]:
    stats = {"typo": 0, "residue": 0, "patterns": 0, "assoc": 0}

    for old, new in TYPO_REPLACEMENTS + RESIDUE_REPLACEMENTS:
        stats["typo" if (old, new) in TYPO_REPLACEMENTS else "residue"] += (
            _replace_text_in_element(doc_root, old, new)
        )
    for pat, repl in TYPO_REGEX_REPLACEMENTS:
        for t in doc_root.iter(qn("w:t")):
            if t.text and pat.search(t.text):
                t.text = pat.sub(repl, t.text)
                stats["typo"] += 1

    for p in doc_root.iter(qn("w:p")):
        text = _paragraph_text(p)
        if not text.strip():
            continue
        new_text = text
        for pat, repl in RESIDUE_LINE_PATTERNS:
            if pat.search(new_text):
                new_text = pat.sub(repl, new_text)
                stats["patterns"] += 1
        # Associations / paragraphes mentorat hors théorie générique
        if CLASS_ASSOC_RE.search(new_text) and (
            "instance de la classe" in new_text.lower()
            or new_text.strip() in {"Validateur.", "Mentor.", "Apprenant."}
        ):
            new_text = (
                "[À COMPLÉTER — association de classes du domaine MOOMOO "
                "(résidu mentorat retiré)]"
            )
            stats["assoc"] += 1
        elif MENTORAT_PARA_RE.search(new_text) and any(
            k in new_text.lower()
            for k in (
                "orientation intelligente",
                "mentorat",
                "mentors",
                "apprenant",
                "validateur",
            )
        ):
            # Ne pas toucher dédicace/remerciements (pas de ces termes)
            # Remplacer mentions résiduelles restantes
            cleaned = new_text
            cleaned = re.sub(
                r"orientation intelligente et de mentorat",
                "[À COMPLÉTER : MOOMOO]",
                cleaned,
                flags=re.I,
            )
            cleaned = re.sub(
                r"\bmentors\b",
                "[À COMPLÉTER]",
                cleaned,
                flags=re.I,
            )
            cleaned = re.sub(
                r"\bvalidateurs\b",
                "[À COMPLÉTER]",
                cleaned,
                flags=re.I,
            )
            cleaned = re.sub(
                r"\bprofessionnels\b",
                "[À COMPLÉTER]",
                cleaned,
                flags=re.I,
            )
            if cleaned != new_text:
                new_text = cleaned
                stats["residue"] += 1

        if new_text != text:
            _rewrite_paragraph_plain(p, new_text)

    return stats


def _rewrite_paragraph_plain(p: etree._Element, text: str, *, italic: bool = False, size_pt: float = FONT_BODY_PT) -> None:
    ppr = _clear_paragraph_content(p)
    # Préserver style
    p.append(_make_run(text, italic=italic, size_pt=size_pt))
    # Remettre pPr en tête
    if ppr is not None and p[0] is not ppr:
        p.remove(ppr)
        p.insert(0, ppr)


def normalize_styles(styles_root: etree._Element) -> None:
    # docDefaults
    rpr_default = styles_root.find(f".//{qn('w:rPrDefault')}/{qn('w:rPr')}")
    if rpr_default is not None:
        _ensure_rfonts(rpr_default)
        _set_sz(rpr_default, FONT_BODY_PT)
    ppr_default = styles_root.find(f".//{qn('w:pPrDefault')}/{qn('w:pPr')}")
    if ppr_default is not None:
        _set_spacing(ppr_default)
        _set_jc(ppr_default, "both")

    for style in styles_root.iter(qn("w:style")):
        sid = style.get(qn("w:styleId")) or ""
        name_el = style.find(qn("w:name"))
        name = name_el.get(qn("w:val")) if name_el is not None else ""
        rpr = style.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            style.append(rpr)
        _ensure_rfonts(rpr)

        ppr = style.find(qn("w:pPr"))
        if ppr is None and style.get(qn("w:type")) == "paragraph":
            ppr = OxmlElement("w:pPr")
            style.insert(0, ppr)

        if sid == "Normal" or name == "Normal":
            _set_sz(rpr, FONT_BODY_PT)
            if ppr is not None:
                _set_jc(ppr, "both")
                _set_spacing(ppr)

        if sid in {"Titre1", "Titre2"} or name in {"heading 1", "heading 2"}:
            if ppr is not None:
                _set_jc(ppr, "center")
            color = rpr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rpr.append(color)
            color.set(qn("w:val"), HEADING1_COLOR)

        if sid == "Titre3" or name == "heading 3":
            color = rpr.find(qn("w:color"))
            if color is None:
                color = OxmlElement("w:color")
                rpr.append(color)
            color.set(qn("w:val"), HEADING3_COLOR)

        # Couleur hors palette documentée → ramener vers 2F5496
        color = rpr.find(qn("w:color"))
        if color is not None:
            val = (color.get(qn("w:val")) or "").upper()
            if val and val not in COLOR_PALETTE and val != "AUTO":
                if val == "1F4E79":
                    color.set(qn("w:val"), HEADING1_COLOR)


def normalize_numbering() -> bytes:
    """Un seul abstractNum : niv1 •, niv2 o (Courier New)."""
    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W_NS}" xmlns:r="{R_NS}">
  <w:abstractNum w:abstractNumId="{LIST_ABSTRACT_ID}">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="{FONT_NAME}" w:hAnsi="{FONT_NAME}" w:cs="{FONT_NAME}" w:eastAsia="{FONT_NAME}"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:eastAsia="Courier New"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="{LIST_NUM_ID}">
    <w:abstractNumId w:val="{LIST_ABSTRACT_ID}"/>
  </w:num>
</w:numbering>
"""
    return xml.encode("utf-8")


def remap_lists(doc_root: etree._Element) -> int:
    n = 0
    for num_id in doc_root.iter(qn("w:numId")):
        num_id.set(qn("w:val"), LIST_NUM_ID)
        n += 1
    # Limiter ilvl à 0/1
    for ilvl in doc_root.iter(qn("w:ilvl")):
        try:
            v = int(ilvl.get(qn("w:val"), "0"))
        except ValueError:
            v = 0
        if v > 1:
            ilvl.set(qn("w:val"), "1")
    return n


def format_paragraphs(doc_root: etree._Element) -> None:
    caption_re = re.compile(r"^(Tableau|Figure)\s*\d*\s*:", re.I)
    source_re = re.compile(r"^Source\s*:", re.I)

    for p in doc_root.iter(qn("w:p")):
        ppr = _get_or_create_ppr(p)
        p_style = ppr.find(qn("w:pStyle"))
        style_id = p_style.get(qn("w:val")) if p_style is not None else ""
        text = _paragraph_text(p).strip()

        # Interligne / justification corps
        if style_id in {"Titre1", "Titre2"}:
            _set_jc(ppr, "center")
        elif style_id.startswith("Titre"):
            pass
        else:
            _set_jc(ppr, "both")
            _set_spacing(ppr, line=LINE_TWIPS, after_pt=SPACING_AFTER_PT)

        # Runs : Times + taille
        is_caption = bool(caption_re.match(text)) or bool(source_re.match(text))
        for r in p.findall(qn("w:r")):
            rpr = _get_or_create_rpr(r)
            _ensure_rfonts(rpr)
            if is_caption:
                _set_sz(rpr, FONT_CAPTION_PT)
                if rpr.find(qn("w:i")) is None:
                    rpr.append(OxmlElement("w:i"))
            elif style_id in {"", "Normal", "Paragraphedeliste"}:
                # Ne forcer 12pt que si pas déjà une taille explicite de titre
                sz = rpr.find(qn("w:sz"))
                if sz is None:
                    _set_sz(rpr, FONT_BODY_PT)


def format_tables(doc_root: etree._Element) -> int:
    count = 0
    for tbl in doc_root.iter(qn("w:tbl")):
        count += 1
        tbl_pr = tbl.find(qn("w:tblPr"))
        if tbl_pr is None:
            tbl_pr = OxmlElement("w:tblPr")
            tbl.insert(0, tbl_pr)

        # Largeur utile
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:w"), str(int(USEFUL_WIDTH_CM * 567)))  # dxa approx
        tbl_w.set(qn("w:type"), "dxa")

        rows = tbl.findall(qn("w:tr"))
        for ri, tr in enumerate(rows):
            tr_pr = tr.find(qn("w:trPr"))
            if tr_pr is None:
                tr_pr = OxmlElement("w:trPr")
                tr.insert(0, tr_pr)
            # cantSplit
            if tr_pr.find(qn("w:cantSplit")) is None:
                tr_pr.append(OxmlElement("w:cantSplit"))
            if ri == 0:
                if tr_pr.find(qn("w:tblHeader")) is None:
                    tr_pr.append(OxmlElement("w:tblHeader"))

            for tc in tr.findall(qn("w:tc")):
                tc_pr = tc.find(qn("w:tcPr"))
                if tc_pr is None:
                    tc_pr = OxmlElement("w:tcPr")
                    tc.insert(0, tc_pr)
                shd = tc_pr.find(qn("w:shd"))
                if shd is None:
                    shd = OxmlElement("w:shd")
                    tc_pr.append(shd)
                fill = TABLE_HEADER_FILL if ri == 0 else TABLE_ALT_FILL
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), fill)

                for p in tc.findall(qn("w:p")):
                    ppr = _get_or_create_ppr(p)
                    _set_spacing(ppr, line=TABLE_LINE_TWIPS, after_pt=0)
                    for r in p.findall(qn("w:r")):
                        rpr = _get_or_create_rpr(r)
                        _ensure_rfonts(rpr)
                        _set_sz(rpr, FONT_TABLE_PT)
                        if ri == 0:
                            # Texte en-tête blanc
                            color = rpr.find(qn("w:color"))
                            if color is None:
                                color = OxmlElement("w:color")
                                rpr.append(color)
                            color.set(qn("w:val"), "FFFFFF")
                            if rpr.find(qn("w:b")) is None:
                                rpr.append(OxmlElement("w:b"))
    return count


def insert_seq_captions(doc_root: etree._Element) -> int:
    """Remplace 'Tableau N: …' / 'Figure N: …' par libellé + champ SEQ."""
    cap_re = re.compile(r"^(Tableau|Figure)\s*(\d+)?\s*:\s*(.*)$", re.I)
    n = 0
    for p in list(doc_root.iter(qn("w:p"))):
        text = _paragraph_text(p).strip()
        m = cap_re.match(text)
        if not m:
            continue
        # Déjà un champ SEQ ?
        instrs = "".join(
            (it.text or "") for it in p.iter(qn("w:instrText"))
        )
        if "SEQ" in instrs.upper():
            continue
        kind = "Tableau" if m.group(1).lower().startswith("tableau") else "Figure"
        rest = m.group(3).strip()
        ppr = _clear_paragraph_content(p)
        _set_jc(ppr, "both")
        _set_spacing(ppr, line=LINE_TWIPS, after_pt=SPACING_AFTER_PT)
        # Style légende si possible — garder Normal + italique 10pt
        p.append(_make_run(f"{kind} ", italic=True, size_pt=FONT_CAPTION_PT))
        for fr in _make_field_runs(f'SEQ {kind} \\* ARABIC', placeholder=m.group(2) or ""):
            p.append(fr)
        if rest:
            p.append(_make_run(f": {rest}", italic=True, size_pt=FONT_CAPTION_PT))
        if ppr is not None:
            p.remove(ppr)
            p.insert(0, ppr)
        n += 1
    return n


def insert_toc_fields(doc_root: etree._Element) -> dict[str, bool]:
    """Insère TOC sous SOMMAIRE / LISTE DES TABLEAUX / LISTE DES FIGURES."""
    targets = {
        "SOMMAIRE": 'TOC \\o "1-3" \\h \\z \\u',
        "LISTE DES TABLEAUX": 'TOC \\h \\z \\c "Tableau"',
        "LISTE DES FIGURES": 'TOC \\h \\z \\c "Figure"',
    }
    done = {k: False for k in targets}
    body = doc_root.find(qn("w:body"))
    if body is None:
        return done

    paragraphs = [c for c in body if c.tag == qn("w:p")]
    for idx, p in enumerate(paragraphs):
        text = _paragraph_text(p).strip().upper()
        for key, instr in targets.items():
            if text == key and not done[key]:
                # Vérifier si le suivant a déjà un TOC
                inserted = False
                if idx + 1 < len(paragraphs):
                    nxt = paragraphs[idx + 1]
                    instrs = "".join(
                        (it.text or "") for it in nxt.iter(qn("w:instrText"))
                    )
                    if "TOC" in instrs.upper():
                        done[key] = True
                        inserted = True
                if not inserted:
                    new_p = OxmlElement("w:p")
                    ppr = _get_or_create_ppr(new_p)
                    _set_spacing(ppr)
                    _set_jc(ppr, "both")
                    for fr in _make_field_runs(
                        instr, placeholder=f"[Mettre à jour le champ : {key}]"
                    ):
                        new_p.append(fr)
                    # Insérer après p
                    parent = p.getparent()
                    parent.insert(parent.index(p) + 1, new_p)
                    done[key] = True
    return done


def update_header(header_xml: bytes) -> bytes:
    root = etree.fromstring(header_xml)
    texts = list(root.iter(qn("w:t")))
    if not texts:
        return header_xml
    # Premier run texte = titre ; vider les doublons
    texts[0].text = HEADER_TITLE
    for t in texts[1:]:
        if t.text and ("CONCEPTION" in t.text.upper() or "MOOMOO" in t.text.upper() or "REALISATION" in t.text.upper() or "RÉALISATION" in t.text.upper()):
            t.text = ""
    _set_all_rfonts(root)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def update_footer(footer_xml: bytes) -> bytes:
    root = etree.fromstring(footer_xml)
    # Conserver texte OTILA ; normaliser année si cassée « 202 5 -202 »
    for t in root.iter(qn("w:t")):
        if t.text and "202" in t.text and "OTILA" not in (t.text or ""):
            # ne pas inventer 2026 si incomplet — signaler via placeholder seulement si residue
            pass
        if t.text:
            for old, new in (("202 5 -202", "2025-2026"), ("202 5-202", "2025-2026")):
                if old in t.text:
                    t.text = t.text.replace(old, new)
    _set_all_rfonts(root)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def ensure_sections_and_margins(doc: Document) -> None:
    for section in doc.sections:
        section.page_width = Cm(PAGE_WIDTH_CM)
        section.page_height = Cm(PAGE_HEIGHT_CM)
        section.left_margin = Cm(MARGIN_CM)
        section.right_margin = Cm(MARGIN_CM)
        section.top_margin = Cm(MARGIN_CM)
        section.bottom_margin = Cm(MARGIN_CM)


def ensure_compat(settings_xml: bytes) -> bytes:
    root = etree.fromstring(settings_xml)
    # Chercher compatibilityMode
    found = False
    for el in root.iter():
        if el.tag.endswith("compatSetting") and el.get(qn("w:name")) == "compatibilityMode":
            el.set(qn("w:val"), COMPAT_MODE)
            found = True
    if not found:
        # Ajouter sous w:compat
        compat = root.find(qn("w:compat"))
        if compat is None:
            compat = OxmlElement("w:compat")
            root.append(compat)
        cs = OxmlElement("w:compatSetting")
        cs.set(qn("w:name"), "compatibilityMode")
        cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
        cs.set(qn("w:val"), COMPAT_MODE)
        compat.append(cs)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def clean_media(parts: dict[str, bytes], zin: zipfile.ZipFile | None = None) -> tuple[set[str], set[str]]:
    """Supprime media non utilisées (document) / wdp/svg/webp.

    Préserve les media encore référencées par header/footer.
    numbering.xml.rels (puces image) est retiré avec le numbering reconstruit.
    Retourne (fichiers_supprimés, parts_rels_à_supprimer).
    """
    doc_xml = parts["word/document.xml"].decode("utf-8")
    used_doc_rids = set(re.findall(r'r:embed="(rId\d+)"', doc_xml))
    used_doc_rids |= set(re.findall(r'r:id="(rId\d+)"', doc_xml))

    # Media encore requises hors document (header/footer)
    protected: set[str] = set()
    for extra, rels_name in (
        ("word/header1.xml", "word/_rels/header1.xml.rels"),
        ("word/footer1.xml", "word/_rels/footer1.xml.rels"),
        ("word/footer2.xml", "word/_rels/footer2.xml.rels"),
    ):
        if extra not in parts:
            continue
        xml = parts[extra].decode("utf-8")
        embeds = set(re.findall(r'r:embed="(rId\d+)"', xml)) | set(
            re.findall(r'r:id="(rId\d+)"', xml)
        )
        if rels_name in parts:
            rel_root = ET.fromstring(parts[rels_name])
            for rel in rel_root:
                tid = rel.attrib.get("Id")
                tgt = rel.attrib.get("Target", "")
                if tid in embeds and "media/" in tgt:
                    protected.add(Path(tgt).name)

    rels_xml = parts["word/_rels/document.xml.rels"]
    root = ET.fromstring(rels_xml)
    removed_files: set[str] = set()
    for rel in list(root):
        tid = rel.attrib.get("Id")
        tgt = rel.attrib.get("Target", "")
        rel_type = rel.attrib.get("Type", "")
        if "media/" not in tgt and "hdphoto" not in rel_type:
            continue
        fname = Path(tgt).name
        lower = fname.lower()
        drop = False
        if lower.endswith((".wdp", ".svg", ".webp")):
            drop = True
        elif tid not in used_doc_rids and fname not in protected:
            drop = True
        if drop and fname not in protected:
            removed_files.add(fname)
            root.remove(rel)

    parts["word/_rels/document.xml.rels"] = ET.tostring(
        root, encoding="utf-8", xml_declaration=True
    )

    # numbering reconstruit sans puces image → supprimer ses rels + gif
    parts_to_delete = {"word/_rels/numbering.xml.rels"}
    if "word/_rels/numbering.xml.rels" in parts:
        num_root = ET.fromstring(parts["word/_rels/numbering.xml.rels"])
        for rel in num_root:
            tgt = rel.attrib.get("Target", "")
            if "media/" in tgt:
                fname = Path(tgt).name
                if fname not in protected:
                    removed_files.add(fname)

    # Ne jamais supprimer les media protégées (header)
    removed_files -= protected
    return removed_files, parts_to_delete


def write_side_docs(stats: dict) -> None:
    a_lines = [
        "# A_VERIFIER — points à compléter (PASSE 1)",
        "",
        "## Bloc utilisateur incomplet (ne pas inventer)",
        "- Superviseur IAI : `[À COMPLÉTER]`",
        "- Maître de stage : `[À COMPLÉTER]` (AZOUTSA retiré — résidu ancien projet)",
        "- Période de stage : `[À COMPLÉTER]` (date 06/08/2026 retirée)",
        "- Nature / description MOOMOO : `[À COMPLÉTER]`",
        "- Langue des signes ciblée : `[À COMPLÉTER]`",
        "- Plateforme (mobile/web) : `[À COMPLÉTER]` — l’en-tête source disait « APPLICATION MOBILE », le titre officiel ne précise pas",
        "- Stack technique retenue : `[À COMPLÉTER]` (mentions React/FastAPI/PostgreSQL/Gemini issues de l’ancien projet — à valider pour MOOMOO)",
        "- Avancement / dataset : `[À COMPLÉTER]`",
        "",
        "## Résidus mentorat remplacés par placeholders",
        "- Cas d’utilisation publications / séances",
        "- Associations de classes Apprenant / Mentor / Validateur / Publication…",
        "- Formulations « orientation intelligente et mentorat », « Nexura »",
        "",
        "## Normalisation (choix PASSE 1)",
        "- Titre officiel appliqué en en-tête (exact)",
        "- Times New Roman forcé (styles + runs + ascii/hAnsi/cs/eastAsia)",
        "- Titre 1 / Titre 2 centrés via styles",
        "- Listes : un seul abstractNum (• / o Courier New)",
        "- Tableaux : en-tête 7030A0, lignes E4CCFC, tblHeader, cantSplit, 11 pt, interligne 1.0",
        "- Champs SEQ Tableau/Figure + TOC sommaire / listes (à mettre à jour dans Word/LibreOffice)",
        "- Media orphelines / non utilisées / WDP supprimées",
        "- Couleur de style 1F4E79 ramenée à 2F5496 (hors palette listée)",
        "",
        f"## Stats script: {stats}",
        "",
    ]
    A_VERIFIER.write_text("\n".join(a_lines), encoding="utf-8")

    p_lines = [
        "# PLACEHOLDERS — inventaire (PASSE 1)",
        "",
        "Les placeholders PARTIE I (`Insérer votre texte ici…`, `NOM DE LA STRUTURE`, etc.) sont **conservés** (contenu hors scope PASSE 1).",
        "",
        "Placeholders introduits / normalisés en PASSE 1 :",
        "- `[À COMPLÉTER]` — métadonnées stagiaire / entreprise / dates",
        "- `[À COMPLÉTER : MOOMOO]` — formulations liées à l’ancien projet",
        "- `[À COMPLÉTER — association de classes du domaine MOOMOO …]`",
        "- `[À COMPLÉTER : cas d’utilisation MOOMOO]` / légendes Figure/Tableau associées",
        "- Champs TOC : `[Mettre à jour le champ : …]` jusqu’à refresh Word",
        "",
        "Rédaction Dossiers 1–7 : passes suivantes.",
        "",
    ]
    PLACEHOLDERS.write_text("\n".join(p_lines), encoding="utf-8")
    print(f"[ok] {A_VERIFIER}")
    print(f"[ok] {PLACEHOLDERS}")


def build() -> dict:
    if not SRC.exists():
        raise SystemExit(f"Source introuvable: {SRC}")

    export_markdown()

    # Copie fraîche (idempotence)
    shutil.copy2(SRC, OUT)

    # Lecture / transformations XML
    with zipfile.ZipFile(OUT, "r") as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    doc_root = etree.fromstring(parts["word/document.xml"])
    styles_root = etree.fromstring(parts["word/styles.xml"])

    stats: dict = {}
    stats["text"] = apply_text_cleanups(doc_root)
    normalize_styles(styles_root)
    stats["lists_remapped"] = remap_lists(doc_root)
    format_paragraphs(doc_root)
    stats["tables"] = format_tables(doc_root)
    stats["seq"] = insert_seq_captions(doc_root)
    stats["toc"] = insert_toc_fields(doc_root)
    _set_all_rfonts(doc_root)

    parts["word/document.xml"] = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    parts["word/styles.xml"] = etree.tostring(
        styles_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    parts["word/numbering.xml"] = normalize_numbering()

    if "word/header1.xml" in parts:
        parts["word/header1.xml"] = update_header(parts["word/header1.xml"])
    for fn in ("word/footer1.xml", "word/footer2.xml"):
        if fn in parts:
            parts[fn] = update_footer(parts[fn])

    if "word/settings.xml" in parts:
        parts["word/settings.xml"] = ensure_compat(parts["word/settings.xml"])

    removed, parts_to_delete = clean_media(zin=None, parts=parts)  # type: ignore[arg-type]
    for pdel in parts_to_delete:
        parts.pop(pdel, None)
    stats["media_removed"] = sorted(removed)

    # Réécriture zip sans media / parts supprimées
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            base = Path(name).name
            if name.startswith("word/media/") and base in removed:
                continue
            if name in parts_to_delete:
                continue
            zout.writestr(name, data)
    OUT.write_bytes(buf.getvalue())

    # Marges via python-docx (sections)
    doc = Document(str(OUT))
    ensure_sections_and_margins(doc)
    # Seconde passe légère runs python-docx (corps)
    for p in doc.paragraphs:
        if p.style and p.style.name in {"Heading 1", "Heading 2"}:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = FONT_NAME
            rpr = run._element.get_or_add_rPr()
            _ensure_rfonts(rpr)
    doc.save(str(OUT))

    write_side_docs(stats)
    print(f"[ok] {OUT} (passe 1)")
    print("stats_passe1:", stats)

    # PASSES 2–4 : contenu, biblio, placeholders diagrammes, side docs
    import sys
    from pathlib import Path as _Path

    _scripts = str(_Path(__file__).resolve().parent)
    if _scripts not in sys.path:
        sys.path.insert(0, _scripts)
    from pass_content import apply_passes

    stats2 = apply_passes(OUT)
    stats["passes_2_4"] = stats2
    print(f"[ok] {OUT} (passes 2–4)")
    print("stats_passes_2_4:", stats2)
    return stats


if __name__ == "__main__":
    build()
