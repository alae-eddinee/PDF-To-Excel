"""
bc_pdf_to_pivot.py — Extraction pivot BON DE COMMANDE (Marjane & LV)
Usage: python bc_pdf_to_pivot.py <fichier.pdf> [output.xlsx]
"""

import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# DÉTECTION DU FORMAT
# ─────────────────────────────────────────────

def detect_format(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    if "MARJANE" in text.upper():
        return "marjane"
    if "HYPER MARCHE LV" in text.upper() or "HYPE MARCHE LV" in text.upper() or "HYPER SUD" in text.upper():
        return "lv"
    if "LIVREA" in text.replace(" ", "").upper():
        return "lv"
    return "marjane"


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def _get_rows(words, y_tolerance=3):
    rows = {}
    for w in words:
        y = round(w['top'] / y_tolerance) * y_tolerance
        rows.setdefault(y, []).append(w)
    return {y: sorted(ws, key=lambda w: w['x0']) for y, ws in sorted(rows.items())}


# ─────────────────────────────────────────────
# PARSEUR MARJANE
# Uses extract_words() with x/y positions to handle multi-column PDFs
# where spaces between columns get merged by extract_text().
#
# Column layout (x positions):
#   Col 1 (x < ~180) : Commande par / MARJANE HOLDING
#   Col 2 (x ~180-350) : Commande a / MEDIDIS
#   Col 3 (x > ~350)  : Livre a / NOM DU MAGASIN
#
# Article row — numbers at the right:
#   nums[-3] = Quant en UC
#   nums[-2] = UVC/UC
#   nums[-1] = Quant en UVC  <- we extract this
# ─────────────────────────────────────────────

def parse_marjane(pdf_path: str) -> tuple[dict, str, str]:
    data = {}
    date_cmd = ""

    EAN_RE = re.compile(r'^\d{13}$')
    DATE_RE = re.compile(r'(\d{2}/\d{2}/\d{2,4})')
    NUM_RE = re.compile(r'^\d+(\.\d+)?$')
    LIVREA_X = 350

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            if not words:
                continue

            rows = _get_rows(words)

            # Date commande
            if not date_cmd:
                for ws in rows.values():
                    row_str = ' '.join(w['text'] for w in ws)
                    m = DATE_RE.search(row_str)
                    if m:
                        date_cmd = m.group(1)
                        break

            # Nom du magasin: ligne contenant MEDIDIS, col 3 (x > LIVREA_X)
            magasin = ""
            for ws in rows.values():
                texts = [w['text'].upper() for w in ws]
                if 'MEDIDIS' in texts:
                    livrea_words = [w for w in ws if w['x0'] > LIVREA_X]
                    if livrea_words:
                        raw = livrea_words[0]['text']
                        raw = re.sub(r'^(MARJANE)([A-Z])', r'\1 \2', raw, flags=re.I)
                        magasin = raw.strip()
                    break

            if not magasin:
                continue

            # Articles: ligne dont le 1er mot est un EAN 13
            for ws in rows.values():
                if not ws:
                    continue
                first = ws[0]['text']
                if not EAN_RE.match(first):
                    continue

                ean = first

                # Libelle: mots non-numeriques apres l'EAN, avant LIVREA_X
                libelle_words = []
                for w in ws[1:]:
                    if w['x0'] >= LIVREA_X:
                        break
                    if not NUM_RE.match(w['text']):
                        libelle_words.append(w['text'])
                    else:
                        break
                libelle = " ".join(libelle_words).strip()

                # nums[-1] = Quant en UVC (rightmost column)
                nums = [w['text'] for w in ws if NUM_RE.match(w['text'])]
                if len(nums) < 3:
                    continue
                qty = float(nums[-1])  # Quant en UVC

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
    data = {}
    date_cmd = ""

    EAN_RE = re.compile(r'\b\d{13}\b')
    DATE_RE = re.compile(r'\b(\d{2}/\d{2}/\d{2})\b')
    ARTICLE_LINE_RE = re.compile(r'^(\d{6})\s+(\d{13})\s+(.+?)\s+Piece\s+\d+\s+\S+\s+\d+\s+([\d.]+)\s*$')
    QTY_END_RE = re.compile(r'([\d.]+)\s*$')

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = [l.strip() for l in text.splitlines() if l.strip()]

            if not date_cmd:
                for line in lines:
                    m = DATE_RE.search(line)
                    if m:
                        date_cmd = m.group(1)
                        break

            magasin = ""
            for i, line in enumerate(lines):
                normalized = line.replace(" ", "").upper()
                if "COMMANDEPARLIVREACOMMANDEA" in normalized:
                    if i + 1 < len(lines):
                        parts = lines[i + 1].split()
                        name_tokens = []
                        for tok in parts:
                            if tok.upper() == "MEDIDIS":
                                break
                            name_tokens.append(tok)
                        half = len(name_tokens) // 2
                        magasin = " ".join(name_tokens[:half] if half > 0 else name_tokens).strip()
                    break

            if not magasin:
                for line in lines:
                    upper = line.upper()
                    if ("HYPER MARCHE LV" in upper or "HYPE MARCHE LV" in upper or "HYPER SUD" in upper):
                        if "MEDIDIS" not in upper and len(line.split()) < 10:
                            magasin = line.strip()
                            break

            if not magasin:
                continue

            in_articles = False
            prev_line = ""
            for line in lines:
                if re.search(r'Code\s*ext|Code\s*EAN|Libelle\s*article', line, re.I):
                    in_articles = True
                    prev_line = ""
                    continue
                if not in_articles:
                    prev_line = line
                    continue
                if re.search(r'Date\s*de\s*livraison|Quantite\s*totale|BON\s*DE\s*COMMANDE', line, re.I):
                    in_articles = False
                    continue

                m = ARTICLE_LINE_RE.match(line)
                if m:
                    ean, libelle_raw, qty = m.group(2), m.group(3).strip(), float(m.group(4))
                    if ean not in data:
                        data[ean] = {"libelle": libelle_raw}
                    data[ean][magasin] = data[ean].get(magasin, 0) + qty
                    prev_line = ""
                    continue

                ean_m = EAN_RE.search(line)
                if not ean_m:
                    prev_line = line
                    continue

                ean = ean_m.group(0)
                qty_m = QTY_END_RE.search(line)
                if not qty_m:
                    prev_line = line
                    continue

                qty = float(qty_m.group(1))
                idx = line.index(ean)
                after_ean = line[idx + len(ean):].strip()
                libelle_parts = [tok for tok in after_ean.split() if not re.match(r'^\d', tok)]
                libelle_raw = " ".join(libelle_parts).strip() or prev_line

                if ean not in data:
                    data[ean] = {"libelle": libelle_raw}
                data[ean][magasin] = data[ean].get(magasin, 0) + qty
                prev_line = ""

    titre = "BON DE COMMANDE — MEDIDIS / LV"
    if date_cmd:
        titre += f" — {date_cmd}"
    return data, date_cmd, titre


# ─────────────────────────────────────────────
# CONSTRUCTION DU PIVOT EXCEL
# Matches reference format exactly:
#   Row 1  : titre merged, dark blue bg, white bold, height 21.95
#   Row 2  : EAN + Libelle (wrap), magasins rotated 90deg, TOTAL header
#            height 165pt, store cols width 3.29 (narrow)
#   Data   : alternating white/EBF3FB, total col D6E4F0 bold
#   Total  : TOTAL GENERAL row, grand total cell dark blue/white
# ─────────────────────────────────────────────

HEADER_BG = "1F4E79"
HEADER_FG = "FFFFFF"
TOTAL_BG  = "D6E4F0"
ALT_BG    = "EBF3FB"

def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, color=color, size=size)


def build_pivot(data: dict, titre: str, output_path: str, fmt: str):
    if not data:
        print("Aucune donnee extraite.")
        return

    magasins = []
    seen = set()
    for row in data.values():
        for k in row:
            if k != "libelle" and k not in seen:
                magasins.append(k)
                seen.add(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "Pivot BC"

    EAN_col       = 1
    LIB_col       = 2
    first_mag_col = 3
    total_col     = 2 + len(magasins) + 1

    # Row 1: Titre
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_col)
    c = ws.cell(1, 1, titre)
    c.font = Font(name="Arial", bold=True, color=HEADER_FG, size=12)
    c.fill = _fill(HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 21.95

    # Row 2: Headers
    for col, label, wrap, rot in [
        (EAN_col, "EAN Article",    True,  0),
        (LIB_col, "Libelle Article", True, 0),
    ]:
        c = ws.cell(2, col, label)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = _border()

    for i, mag in enumerate(magasins):
        c = ws.cell(2, first_mag_col + i, mag)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
        c.border = _border()

    c = ws.cell(2, total_col, "TOTAL GENERAL")
    c.font = _font(bold=True, color=HEADER_FG)
    c.fill = _fill(HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = _border()

    ws.row_dimensions[2].height = 165.0

    # Data rows
    rows_written = []
    for row_idx, (ean, row_data) in enumerate(data.items(), 3):
        bg = ALT_BG if (row_idx % 2 == 0) else "FFFFFF"

        c = ws.cell(row_idx, EAN_col, str(ean))
        c.font = _font(); c.fill = _fill(bg); c.border = _border()

        c = ws.cell(row_idx, LIB_col, row_data.get("libelle", ""))
        c.font = _font(); c.fill = _fill(bg); c.border = _border()

        for i, mag in enumerate(magasins):
            qty = row_data.get(mag, None)
            c = ws.cell(row_idx, first_mag_col + i, qty)
            c.font = _font(); c.fill = _fill(bg); c.border = _border()
            if qty is not None:
                c.number_format = "# ###"
                c.alignment = Alignment(horizontal="center")

        col_s = get_column_letter(first_mag_col)
        col_e = get_column_letter(first_mag_col + len(magasins) - 1)
        c = ws.cell(row_idx, total_col, f"=SUM({col_s}{row_idx}:{col_e}{row_idx})")
        c.font = _font(bold=True); c.fill = _fill(TOTAL_BG); c.border = _border()
        c.number_format = "# ###"; c.alignment = Alignment(horizontal="center")
        rows_written.append(row_idx)

    # TOTAL GENERAL row
    if rows_written:
        tr = rows_written[-1] + 1
        r1, r2 = rows_written[0], rows_written[-1]

        ws.merge_cells(start_row=tr, start_column=EAN_col, end_row=tr, end_column=LIB_col)
        c = ws.cell(tr, EAN_col, "TOTAL GENERAL")
        c.font = _font(bold=True); c.fill = _fill(TOTAL_BG); c.border = _border()

        for i in range(len(magasins)):
            col = first_mag_col + i
            cl = get_column_letter(col)
            c = ws.cell(tr, col, f"=SUM({cl}{r1}:{cl}{r2})")
            c.font = _font(bold=True); c.fill = _fill(TOTAL_BG); c.border = _border()
            c.number_format = "# ###"; c.alignment = Alignment(horizontal="center")

        tcl = get_column_letter(total_col)
        c = ws.cell(tr, total_col, f"=SUM({tcl}{r1}:{tcl}{r2})")
        c.font = Font(name="Arial", bold=True, color=HEADER_FG, size=10)
        c.fill = _fill(HEADER_BG); c.border = _border()
        c.number_format = "# ###"; c.alignment = Alignment(horizontal="center")

    # Column widths
    ws.column_dimensions[get_column_letter(EAN_col)].width = 14.14
    ws.column_dimensions[get_column_letter(LIB_col)].width = 31.29
    for i in range(len(magasins)):
        ws.column_dimensions[get_column_letter(first_mag_col + i)].width = 3.29
    ws.column_dimensions[get_column_letter(total_col)].width = 13.0

    ws.freeze_panes = ws.cell(3, first_mag_col)

    wb.save(output_path)
    print(f"Pivot genere : {output_path} | Format: {fmt.upper()} | Articles: {len(data)} | Magasins: {len(magasins)}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def process_pdf(pdf_path: str, output_path: str):
    if not Path(pdf_path).exists():
        print(f"Fichier introuvable : {pdf_path}")
        return
    fmt = detect_format(pdf_path)
    print(f"Format detecte : {fmt.upper()}")
    data, date_cmd, titre = parse_marjane(pdf_path) if fmt == "marjane" else parse_lv(pdf_path)
    build_pivot(data, titre, output_path, fmt)


def main():
    if len(sys.argv) >= 2:
        pdf_path = sys.argv[1]
        output_path = sys.argv[2] if len(sys.argv) >= 3 else f"pivot_{Path(pdf_path).stem}.xlsx"
        process_pdf(pdf_path, output_path)
        return

    pdf_files = list(Path(".").glob("*.pdf"))
    if not pdf_files:
        print("Aucun fichier PDF trouve. Usage: python bc_pdf_to_pivot.py <fichier.pdf> [output.xlsx]")
        sys.exit(1)

    for pdf_file in pdf_files:
        print(f"\n{'='*50}\nTraitement : {pdf_file.name}")
        process_pdf(str(pdf_file), f"pivot_{pdf_file.stem}.xlsx")


if __name__ == "__main__":
    main()
