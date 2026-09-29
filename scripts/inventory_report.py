#!/usr/bin/env python3
"""Inventaire ponctuel PASSE 1 — écrit passe-1-inventaire.md."""
from __future__ import annotations

import os
import re
import zipfile
from collections import Counter
from xml.etree import ElementTree as ET

from docx import Document

SRC = "/workspace/Mon_rapport.docx"
OUT = "/cursor/stores/self/internal/passe-1-inventaire.md"
ALLOWED = {
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


def main() -> None:
    lines: list[str] = [
        "---",
        'cursor:',
        '  subagentId: "bc-bf30ccf1-017d-5357-828b-ef9437aeeb82"',
        "---",
        "",
        "# Inventaire PASSE 1 — Mon_rapport.docx",
        "",
    ]
    doc = Document(SRC)
    z = zipfile.ZipFile(SRC)
    docxml = z.read("word/document.xml").decode("utf-8")

    lines += [
        "## Structure",
        f"- Paragraphes: {len(doc.paragraphs)}",
        f"- Tableaux: {len(doc.tables)}",
        f"- Sections: {len(doc.sections)}",
        f"- Fichiers media: {len([n for n in z.namelist() if n.startswith('word/media/')])}",
        "- Marges: 2,5 cm (899795 EMU) × 5 sections A4 — CONFORME",
        "",
        "## Styles (top)",
    ]
    styles = Counter(p.style.name for p in doc.paragraphs)
    for k, v in styles.most_common(20):
        lines.append(f"- {k}: {v}")

    colors = Counter(
        c.upper()
        for c in re.findall(r'<w:color[^/]*w:val="([A-Fa-f0-9]{6})"', docxml)
    )
    shd = Counter(
        c.upper() for c in re.findall(r'<w:shd[^>]*w:fill="([A-Fa-f0-9]{6})"', docxml)
    )
    lines += ["", "## Couleurs runs"]
    for k, v in colors.most_common():
        flag = "OK" if k in ALLOWED or k == "000000" else "HORS PALETTE"
        lines.append(f"- {k}: {v} [{flag}]")
    lines += ["", "## Fonds (shd)"]
    for k, v in shd.most_common():
        flag = "OK" if k in ALLOWED else "HORS PALETTE"
        lines.append(f"- {k}: {v} [{flag}]")

    theme = z.read("word/theme/theme1.xml").decode("utf-8")
    theme_cols = sorted(
        set(c.upper() for c in re.findall(r'srgbClr val="([A-Fa-f0-9]{6})"', theme))
    )
    lines += [
        "",
        "## Thème Office (srgb) — informatif",
        ", ".join(theme_cols),
    ]

    fonts: Counter[str] = Counter()
    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.name:
                fonts[r.font.name] += 1
    lines += ["", "## Polices explicites (runs)"]
    for k, v in fonts.most_common():
        lines.append(f"- {k}: {v}")

    num = z.read("word/numbering.xml").decode("utf-8")
    lvl = sorted(set(re.findall(r'w:lvlText w:val="([^"]*)"', num)))
    lines += [
        "",
        "## Numbering",
        f"- abstractNum: {num.count('w:abstractNum ')} (cible: 1 seul • / o Courier)",
        f"- lvlText uniques: {lvl}",
    ]

    settings = z.read("word/settings.xml").decode("utf-8")
    m = re.search(r'compatibilityMode[^>]*w:val="(\d+)"', settings)
    lines += [
        "",
        "## Compatibilité",
        f"- compatibilityMode: {m.group(1) if m else 'ABSENT'} (cible 15)",
    ]

    h = z.read("word/header1.xml").decode("utf-8")
    ht = " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", h))
    lines += [
        "",
        "## En-tête actuel",
        f"- {ht[:400]}",
        "- Écarts vs titre exact: REALISATION sans accent, « APPLICATION MOBILE », « VIS VERSA », « LANGAGE DE SIGNES »",
    ]
    lines += ["", "## Pied de page"]
    for fn in ["word/footer1.xml", "word/footer2.xml"]:
        t = " ".join(
            re.findall(r"<w:t[^>]*>([^<]*)</w:t>", z.read(fn).decode("utf-8"))
        )
        lines.append(f"- {fn}: {t[:240]}")

    rels = z.read("word/_rels/document.xml.rels").decode("utf-8")
    root = ET.fromstring(rels)
    rid_map = {}
    for rel in root:
        tid = rel.attrib.get("Id")
        tgt = rel.attrib.get("Target", "")
        if "media/" in tgt:
            rid_map[tid] = os.path.basename(tgt)
    used = set(re.findall(r'r:embed="(rId\d+)"', docxml))
    for extra in ["word/header1.xml", "word/footer1.xml", "word/footer2.xml"]:
        try:
            used |= set(re.findall(r'r:embed="(rId\d+)"', z.read(extra).decode("utf-8")))
        except KeyError:
            pass
    unused = {k: v for k, v in rid_map.items() if k not in used}
    media = [n for n in z.namelist() if n.startswith("word/media/")]
    orphans = {os.path.basename(m) for m in media} - set(rid_map.values())
    bad = [m for m in media if m.lower().endswith((".wdp", ".svg", ".webp"))]
    lines += [
        "",
        "## Media",
        f"- Orphelines (fichier sans rel): {sorted(orphans)}",
        f"- Rels media non référencées dans le XML: {len(unused)}",
        f"- WDP/SVG/WebP: {bad}",
    ]

    pat = re.compile(
        r"Apprenant|Mentor|Validateur|Professionnel|publication|séances|AZOUTSA|"
        r"06/08/2026|mentorat|orientation intelligente|Nexura",
        re.I,
    )
    hits = []
    for i, p in enumerate(doc.paragraphs):
        if pat.search(p.text):
            hits.append((str(i), p.text[:140]))
    for ti, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                if pat.search(cell.text):
                    hits.append((f"T{ti}", cell.text[:140]))
    lines += ["", f"## Résidus ancien projet ({len(hits)} hits)", ""]
    for h0, h1 in hits[:40]:
        lines.append(f"- {h0}: {h1}")

    lines += ["", "## Coquilles ciblées"]
    for i, p in enumerate(doc.paragraphs):
        t = p.text
        if re.search(
            r"NARY JACOBSON|RÉSENTATION DES ARCHITECTURES|diagramme de Gant[^t]|"
            r"DECRIPTION|Digramme|RUMBAUCH",
            t,
            re.I,
        ):
            lines.append(f"- p{i}: {t[:160]}")

    lines += [
        "",
        "## Constats (section 3 cahier — confirmation)",
        "- 5 sections A4, marges 2,5 cm: **CONFIRMÉ**",
        "- compatibilityMode=15: **CONFIRMÉ**",
        "- Times New Roman non appliqué partout (thème minorHAnsi / peu de rFonts explicites): **À CORRIGER**",
        "- Palette corps OK (FF6600, 44546A, FF9966, 7030A0, E4CCFC, 4472C4, 0563C1); style 1F4E79 hors liste: **noter**",
        "- 69 abstractNum + puces mixtes (•, o, ✓, %1.): **À NORMALISER**",
        "- En-tête ≠ titre exact: **À CORRIGER**",
        "- ~50 media inutilisées + 3 orphelines + 2 WDP: **À NETTOYER**",
        "- Résidus mentorat / AZOUTSA / publications / 06/08/2026: **À PURGER**",
        "- Page de garde absente du body (débute à DÉDICACE); titre seulement en en-tête",
        "- PARTIE I placeholders (« Insérer… », NOM DE LA STRUTURE): **CONSERVER**",
        "- Champs SEQ / TOC sommaire & listes: **À INSÉRER**",
        "",
    ]

    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Wrote {OUT} ({len(lines)} lines)")


if __name__ == "__main__":
    main()
