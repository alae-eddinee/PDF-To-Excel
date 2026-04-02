"""
bc_pdf_to_pivot.py — Extraction pivot BON DE COMMANDE (Marjane & LV)
Usage: python bc_pdf_to_pivot.py <fichier.pdf> [output.xlsx]
"""

import re
import sys
from pathlib import Path
from datetime import datetime

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# DÉTECTION DU FORMAT
# ─────────────────────────────────────────────

def detect_format(pdf_path: str) -> str:
    """Retourne 'marjane' ou 'lv' selon le contenu du PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    if "MARJANE" in text.upper():
        return "marjane"
    if "HYPER MARCHE LV" in text.upper() or "HYPE MARCHE LV" in text.upper() or "HYPER SUD" in text.upper():
        return "lv"
    # fallback : LV si on voit "Livre a" sans Marjane
    if "LIVREA" in text.replace(" ", "").upper():
        return "lv"
    return "marjane"


# ─────────────────────────────────────────────
# PARSEUR MARJANE (logique originale)
# ─────────────────────────────────────────────

def parse_marjane(pdf_path: str) -> tuple[dict, str, str]:
    """
    Retourne (data, date_cmd, titre)
    data = { ean: {libelle, magasin: qty, ...} }
    """
    data = {}
    date_cmd = ""

    EAN_RE = re.compile(r'^\d{13}$')
    DATE_RE = re.compile(r'(\d{2}/\d{2}/\d{2,4})')
    QTY_RE = re.compile(r'(\d+(?:\.\d+)?)\s*$')

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = [l.strip() for l in text.splitlines() if l.strip()]

            # Date commande
            if not date_cmd:
                for line in lines:
                    m = DATE_RE.search(line)
                    if m:
                        date_cmd = m.group(1)
                        break

            # Nom du magasin : ligne "Livré à" ou "MARJANE ..."
            magasin = ""
            for i, line in enumerate(lines):
                upper = line.upper()
                if "MARJANE" in upper and not any(
                    kw in upper for kw in ["HOLDING", "MEDIDIS", "COMMANDE", "BON DE"]
                ):
                    # Prendre le premier token MARJANE xxx
                    m = re.search(r'(MARJANE\s+\S+(?:\s+\S+)*)', line, re.I)
                    if m:
                        magasin = m.group(1).strip()
                        break

            if not magasin:
                continue

            # Articles : lignes avec EAN 13 chiffres
            for i, line in enumerate(lines):
                parts = line.split()
                if not parts:
                    continue
                ean = None
                # Chercher un EAN dans la ligne
                for p in parts:
                    if EAN_RE.match(p):
                        ean = p
                        break
                if not ean:
                    continue

                # Quantité : dernier nombre de la ligne
                m = QTY_RE.search(line)
                if not m:
                    continue
                qty = float(m.group(1))

                # Libellé : tout ce qui est entre EAN et le reste
                idx = line.index(ean)
                after_ean = line[idx + len(ean):].strip()
                # Retirer la quantité finale et le reste numérique
                libelle_parts = []
                for tok in after_ean.split():
                    if re.match(r'^\d+(?:\.\d+)?$', tok):
                        break
                    libelle_parts.append(tok)
                libelle = " ".join(libelle_parts).strip()

                if ean not in data:
                    data[ean] = {"libelle": libelle}
                data[ean][magasin] = data[ean].get(magasin, 0) + qty

    titre = "BON DE COMMANDE — MEDIDIS / MARJANE HOLDING"
    if date_cmd:
        titre += f" — {date_cmd}"
    return data, date_cmd, titre


# ─────────────────────────────────────────────
# PARSEUR LV
# ─────────────────────────────────────────────

def parse_lv(pdf_path: str) -> tuple[dict, str, str]:
    """
    Retourne (data, date_cmd, titre)
    data = { ean: {libelle, magasin: qty, ...} }

    Structure LV (par page) :
    - No commande  Date commande ...
    - 2603054920224  18/03/26 12:22  1157 ...
    - Commande par  Livre a  Commande a
    - HYPER MARCHE LV SALE   HYPER MARCHE LV SALE   MEDIDIS
    - ...lignes adresse...
    - Code externe  Code EAN  Libelle article  Type U.C. ...
    - 751430  2000003398768  PLAT A FOUR RECT ...  Piece  0  ...  1  120.000
    - (ligne dimensions)
    - ...
    - Date de livraison souhaitee  ...  Quantite totale
    - 07/05/26 07:00  ...  480.000
    """
    data = {}
    date_cmd = ""

    EAN_RE = re.compile(r'\b\d{13}\b')
    DATE_RE = re.compile(r'\b(\d{2}/\d{2}/\d{2})\b')
    # Article line: starts with 6-digit code externe, then EAN
    ARTICLE_LINE_RE = re.compile(r'^(\d{6})\s+(\d{13})\s+(.+?)\s+Piece\s+\d+\s+\S+\s+\d+\s+([\d.]+)\s*$')
    # Fallback for article lines without full match
    QTY_END_RE = re.compile(r'([\d.]+)\s*$')

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = [l.strip() for l in text.splitlines() if l.strip()]

            # ── Date commande ──
            if not date_cmd:
                for line in lines:
                    m = DATE_RE.search(line)
                    if m:
                        date_cmd = m.group(1)
                        break

            # ── Nom du magasin ──
            # Trouver "Commandepar Livrea Commandea" puis prendre le 1er token de la ligne suivante
            magasin = ""
            for i, line in enumerate(lines):
                normalized = line.replace(" ", "").upper()
                if "COMMANDEPARLIVREACOMMANDEA" in normalized:
                    # La ligne suivante contient : NOM_MAG   NOM_MAG   MEDIDIS
                    if i + 1 < len(lines):
                        next_line = lines[i + 1]
                        # Le magasin est répété deux fois suivi de MEDIDIS
                        # On prend tout jusqu'au 2e bloc répété ou jusqu'à MEDIDIS
                        # Stratégie : diviser en tokens et prendre jusqu'à la répétition
                        parts = next_line.split()
                        # Chercher où le nom se répète ou où MEDIDIS apparaît
                        half = len(parts) // 2
                        # Le nom est la 1ère moitié (avant MEDIDIS)
                        # Heuristique : prendre les tokens jusqu'à "MEDIDIS"
                        name_tokens = []
                        for tok in parts:
                            if tok.upper() == "MEDIDIS":
                                break
                            name_tokens.append(tok)
                        # La 2ème moitié est la répétition → prendre la 1ère moitié
                        half_count = len(name_tokens) // 2
                        if half_count > 0:
                            magasin = " ".join(name_tokens[:half_count])
                        else:
                            magasin = " ".join(name_tokens)
                        magasin = magasin.strip()
                    break

            if not magasin:
                # Fallback : chercher "HYPER MARCHE LV" ou "HYPER SUD"
                for line in lines:
                    upper = line.upper()
                    if ("HYPER MARCHE LV" in upper or "HYPE MARCHE LV" in upper or
                            "HYPER SUD" in upper):
                        # Éviter lignes contenant MEDIDIS ou adresses
                        if "MEDIDIS" not in upper and len(line.split()) < 10:
                            magasin = line.strip()
                            break

            if not magasin:
                continue

            # ── Articles ──
            # Trouver la section articles (après la ligne d'en-tête "Code externe Code EAN ...")
            in_articles = False
            prev_line = ""
            for line in lines:
                # Début section articles
                if "CODEEAN" in line.replace(" ", "").upper() and "LIBELLEARTICLE" in line.replace(" ", "").upper():
                    in_articles = True
                    continue

                # Fin section articles
                if in_articles and ("DATEDELIVRAISON" in line.replace(" ", "").upper()
                                    or "QUANTITETOTALE" in line.replace(" ", "").upper()):
                    in_articles = False
                    break

                if not in_articles:
                    continue

                # Ligne article complète : commence par code externe 6 chiffres
                if re.match(r'^\d{6}\s', line):
                    m = ARTICLE_LINE_RE.match(line)
                    if m:
                        ean = m.group(2)
                        libelle_raw = m.group(3).strip()
                        qty = float(m.group(4))
                    else:
                        # Parsing manuel
                        parts = line.split()
                        ean = None
                        for p in parts:
                            if re.match(r'^\d{13}$', p):
                                ean = p
                                break
                        if not ean:
                            prev_line = line
                            continue
                        # Quantité = dernier nombre
                        qm = QTY_END_RE.search(line)
                        if not qm:
                            prev_line = line
                            continue
                        qty = float(qm.group(1))
                        # Libellé : entre EAN et "Piece"
                        idx_ean = line.index(ean) + len(ean)
                        after = line[idx_ean:].strip()
                        libelle_raw = re.split(r'\s+Piece\s+', after, maxsplit=1)[0].strip()

                    # La ligne suivante contient les dimensions → on les ajoute au libellé
                    prev_line = (ean, libelle_raw, qty)

                elif prev_line and isinstance(prev_line, tuple):
                    # Ligne dimensions (ex: "39X22X8CM")
                    ean, libelle_raw, qty = prev_line
                    dim_candidate = line.strip()
                    # Vérifier que c'est bien des dimensions (contient X et chiffres, pas de EAN)
                    if re.match(r'^[\dX\-\s,]+CM$', dim_candidate, re.I) or re.match(r'^\d+X\d+', dim_candidate):
                        libelle = f"{libelle_raw} {dim_candidate}"
                    else:
                        libelle = libelle_raw

                    if ean not in data:
                        data[ean] = {"libelle": libelle}
                    data[ean][magasin] = data[ean].get(magasin, 0) + qty
                    prev_line = ""
                else:
                    prev_line = line

            # Si la dernière ligne article n'avait pas de ligne dimension suivante
            if isinstance(prev_line, tuple):
                ean, libelle_raw, qty = prev_line
                if ean not in data:
                    data[ean] = {"libelle": libelle_raw}
                data[ean][magasin] = data[ean].get(magasin, 0) + qty

    titre = "BON DE COMMANDE — MEDIDIS / LV"
    if date_cmd:
        titre += f" — {date_cmd}"
    return data, date_cmd, titre


# ─────────────────────────────────────────────
# CONSTRUCTION DU PIVOT EXCEL
# ─────────────────────────────────────────────

HEADER_BG   = "1F4E79"   # bleu foncé
HEADER_FG   = "FFFFFF"
TOTAL_BG    = "D6E4F0"   # bleu clair pour totaux
SUBHDR_BG   = "BDD7EE"
ALT_BG      = "EBF3FB"

def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

def _font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)


def build_pivot(data: dict, titre: str, output_path: str, fmt: str):
    if not data:
        print("⚠ Aucune donnée extraite.")
        return

    # Collecter tous les magasins (ordre d'apparition)
    magasins = []
    seen = set()
    for ean, row in data.items():
        for k in row:
            if k != "libelle" and k not in seen:
                magasins.append(k)
                seen.add(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot BC"

    # ── Titre ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + len(magasins) + 1)
    title_cell = ws.cell(1, 1, titre)
    title_cell.font = Font(name="Arial", bold=True, color=HEADER_FG, size=12)
    title_cell.fill = _fill(HEADER_BG)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── En-têtes ──
    headers = ["EAN Article", "Libellé Article"] + magasins + ["TOTAL GÉNÉRAL"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(2, col, h)
        c.font = _font(bold=True, color=HEADER_FG, size=10)
        c.fill = _fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _border()
    ws.row_dimensions[2].height = 30

    # ── Données ──
    EAN_col = 1
    LIB_col = 2
    first_mag_col = 3
    total_col = 2 + len(magasins) + 1

    rows_written = []
    for row_idx, (ean, row_data) in enumerate(data.items(), 3):
        alt = (row_idx % 2 == 0)
        bg = ALT_BG if alt else "FFFFFF"

        ws.cell(row_idx, EAN_col, str(ean)).border = _border()
        ws.cell(row_idx, EAN_col).font = _font()
        ws.cell(row_idx, EAN_col).fill = _fill(bg)

        ws.cell(row_idx, LIB_col, row_data.get("libelle", "")).border = _border()
        ws.cell(row_idx, LIB_col).font = _font()
        ws.cell(row_idx, LIB_col).fill = _fill(bg)

        for mag_idx, mag in enumerate(magasins):
            col = first_mag_col + mag_idx
            qty = row_data.get(mag, None)
            c = ws.cell(row_idx, col, qty)
            c.border = _border()
            c.font = _font()
            c.fill = _fill(bg)
            if qty is not None:
                c.number_format = "#,##0"
                c.alignment = Alignment(horizontal="center")

        # Total ligne
        col_start = get_column_letter(first_mag_col)
        col_end = get_column_letter(first_mag_col + len(magasins) - 1)
        tc = ws.cell(row_idx, total_col, f"=SUM({col_start}{row_idx}:{col_end}{row_idx})")
        tc.border = _border()
        tc.font = _font(bold=True)
        tc.fill = _fill(TOTAL_BG)
        tc.number_format = "#,##0"
        tc.alignment = Alignment(horizontal="center")
        rows_written.append(row_idx)

    # ── Ligne TOTAL GÉNÉRAL ──
    if rows_written:
        total_row = rows_written[-1] + 1
        ws.cell(total_row, EAN_col, "TOTAL GÉNÉRAL").font = _font(bold=True)
        ws.cell(total_row, EAN_col).fill = _fill(TOTAL_BG)
        ws.cell(total_row, EAN_col).border = _border()
        ws.merge_cells(start_row=total_row, start_column=EAN_col,
                       end_row=total_row, end_column=LIB_col)

        for mag_idx in range(len(magasins)):
            col = first_mag_col + mag_idx
            col_letter = get_column_letter(col)
            r1, r2 = rows_written[0], rows_written[-1]
            c = ws.cell(total_row, col, f"=SUM({col_letter}{r1}:{col_letter}{r2})")
            c.font = _font(bold=True)
            c.fill = _fill(TOTAL_BG)
            c.border = _border()
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="center")

        # Total général de la dernière colonne
        tc_letter = get_column_letter(total_col)
        gt = ws.cell(total_row, total_col,
                     f"=SUM({tc_letter}{rows_written[0]}:{tc_letter}{rows_written[-1]})")
        gt.font = _font(bold=True)
        gt.fill = _fill(HEADER_BG)
        gt.font = Font(name="Arial", bold=True, color=HEADER_FG, size=10)
        gt.border = _border()
        gt.number_format = "#,##0"
        gt.alignment = Alignment(horizontal="center")

    # ── Largeurs colonnes ──
    ws.column_dimensions[get_column_letter(EAN_col)].width = 16
    ws.column_dimensions[get_column_letter(LIB_col)].width = 40
    for i in range(len(magasins)):
        ws.column_dimensions[get_column_letter(first_mag_col + i)].width = 18
    ws.column_dimensions[get_column_letter(total_col)].width = 16

    # ── Figer les volets ──
    ws.freeze_panes = ws.cell(3, first_mag_col)

    try:
        wb.save(output_path)
        print(f"✅ Pivot généré : {output_path}")
        print(f"   Format détecté : {fmt.upper()}")
        print(f"   Articles : {len(data)}")
        print(f"   Magasins : {len(magasins)}")
    except Exception as e:
        print(f"❌ Erreur lors de la sauvegarde : {e}")
        raise


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    # Check for command line arguments first
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
        if not Path(pdf_path).exists():
            print(f"Fichier introuvable : {pdf_path}")
            sys.exit(1)
        
        # Output par défaut
        if len(sys.argv) >= 3:
            output_path = sys.argv[2]
        else:
            stem = Path(pdf_path).stem
            output_path = f"pivot_{stem}.xlsx"
        
        process_pdf(pdf_path, output_path)
        return
    
    # Auto-detect PDF files in current directory
    current_dir = Path(".")
    pdf_files = list(current_dir.glob("*.pdf"))
    
    if not pdf_files:
        print("❌ Aucun fichier PDF trouvé dans le répertoire courant.")
        print("Usage: python bc_pdf_to_pivot.py <fichier.pdf> [output.xlsx]")
        sys.exit(1)
    
    print(f"📁 {len(pdf_files)} fichier(s) PDF trouvé(s) :")
    for i, pdf_file in enumerate(pdf_files, 1):
        print(f"   {i}. {pdf_file.name}")
    
    # Process each PDF file
    for pdf_file in pdf_files:
        stem = pdf_file.stem
        output_path = f"pivot_{stem}.xlsx"
        print(f"\n" + "="*50)
        print(f"Traitement de : {pdf_file.name}")
        print("="*50)
        process_pdf(str(pdf_file), output_path)

def process_pdf(pdf_path: str, output_path: str):
    """Process a single PDF file and generate pivot table"""
    if not Path(pdf_path).exists():
        print(f"Fichier introuvable : {pdf_path}")
        return
    
    fmt = detect_format(pdf_path)
    print(f"📄 Format détecté : {fmt.upper()}")

    if fmt == "marjane":
        data, date_cmd, titre = parse_marjane(pdf_path)
    else:
        data, date_cmd, titre = parse_lv(pdf_path)

    build_pivot(data, titre, output_path, fmt)


if __name__ == "__main__":
    main()