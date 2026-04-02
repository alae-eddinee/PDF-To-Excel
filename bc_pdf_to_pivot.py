"""
bc_pdf_to_pivot.py — Extraction pivot BON DE COMMANDE (Marjane, Marjane-MEDIDIS & LV)
Usage: python bc_pdf_to_pivot.py <fichier.pdf> [output.xlsx]

Formats supportes:
  - marjane  : BON DE COMMANDE Marjane classique (colonnes Livre a / MEDIDIS separees)
  - medidis  : BON DE COMMANDE MEDIDIS / MARJANE HOLDING (colonne unique, libelle sur 2 lignes)
  - lv       : BON DE COMMANDE Hyper Marché LV
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
    upper = text.upper()
    if "HYPER MARCHE LV" in upper or "HYPE MARCHE LV" in upper or "HYPER SUD" in upper:
        return "lv"
    if "LIVREA" in text.replace(" ", "").upper():
        # MEDIDIS format: "Commande par" / "Livre a" / "Commande a" header exists,
        # but the store name is in the "Commande par" column (x < 180) on the MEDIDIS row.
        # Distinguish from classic Marjane by checking the No.ligne column header.
        if "Noligne" in text.replace(" ", "") or "No ligne" in text or "Noligne" in text:
            return "medidis"
        return "marjane"
    if "MARJANE" in upper:
        # Could be either marjane or medidis — inspect column headers
        if "Noligne" in text.replace(" ", "") or "No.ligne" in text.replace(" ", ""):
            return "medidis"
        return "marjane"
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


def _normalize_magasin(raw: str) -> str:
    """Fix merged store names: MARJANEBOUREGREG -> MARJANE BOUREGREG"""
    s = re.sub(r'^(MARJANE)([A-Z])', r'\1 \2', raw, flags=re.I)
    return s.strip()


def _normalize_libelle(text: str) -> str:
    """Fix merged libellé words: CADREPHOTOZEYNA- -> CADRE PHOTO ZEYNA-"""
    s = re.sub(r'(CADRE)(PHOTO|MYLARD)', r'\1 \2', text, flags=re.I)
    s = re.sub(r'(PHOTO)(ZEYNA|LEA|GULIA|RITA|FLECHE)', r'\1 \2', s, flags=re.I)
    return s.strip()


# ─────────────────────────────────────────────
# PARSEUR MEDIDIS
#
# PDF structure (x positions):
#   x ~31  : Commande par  → NOM DU MAGASIN (e.g. MARJANEBOUREGREG)
#   x ~208 : Livre a       → same store name repeated
#   x ~386 : Commande a    → MEDIDIS (the supplier)
#
# Article rows span 2 y-levels:
#   Line 1: [30]EAN [79]LIBELLE_PART1 [160]VL [170]NO_LIGNE [226]PCB
#            [263]Quant_UC [293]UVC/UC [324]Quant_UVC [357]OSSAGA3
#   Line 2: [79]LIBELLE_PART2  (e.g. "10X15CM-BLANC")
#
# EAN can be 12 or 13 digits (some articles use 12-digit codes).
# No.ligne is always 13 digits starting with 078... — must not be confused with EAN.
# Quant en UVC = nums[-1] in the data row (rightmost numeric token).
#
# Key challenge: No.ligne (e.g. 0784313020021) is 13 digits → collides with EAN regex.
# We disambiguate by x-position: EAN is at x < 60, No.ligne is at x ~170.
# ─────────────────────────────────────────────

def parse_medidis(pdf_path: str) -> tuple[dict, str, str]:
    data = {}
    date_cmd = ""

    EAN_RE  = re.compile(r'^\d{12,13}$')
    DATE_RE = re.compile(r'(\d{2}/\d{2}/\d{2,4})')
    NUM_RE  = re.compile(r'^\d+(\.\d+)?$')

    # x-position thresholds (from word-position inspection)
    EAN_X_MAX    = 60    # EAN starts at x~30-38
    LIBELLE_X    = 79    # libellé starts at x~79
    NO_LIGNE_X   = 160   # No.ligne column starts at x~160
    QTY_UVC_X    = 310   # Quant en UVC column at x~324; use 310 as left boundary

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            if not words:
                continue

            rows = _get_rows(words)
            sorted_ys = sorted(rows.keys())

            # Date commande
            if not date_cmd:
                for ws in rows.values():
                    row_str = ' '.join(w['text'] for w in ws)
                    m = DATE_RE.search(row_str)
                    if m and '/' in m.group(1):
                        date_cmd = m.group(1)
                        break

            # Magasin: row where MEDIDIS appears at x~386 → take word(s) at x < 180
            magasin = ""
            for ws in rows.values():
                texts_upper = [w['text'].upper() for w in ws]
                if 'MEDIDIS' in texts_upper:
                    store_words = [w for w in ws if w['x0'] < 180]
                    if store_words:
                        magasin = _normalize_magasin(store_words[0]['text'])
                    break

            if not magasin:
                continue

            # Build a lookup from y → next_y for libellé continuation
            y_list = sorted_ys

            # Parse article rows: EAN at x < EAN_X_MAX
            for idx, y in enumerate(y_list):
                ws = rows[y]
                if not ws:
                    continue

                # First word must be EAN-like and positioned on the left
                first = ws[0]
                if first['x0'] > EAN_X_MAX:
                    continue
                if not EAN_RE.match(first['text']):
                    continue

                ean = first['text']

                # Libellé part 1: words at x~79 up to NO_LIGNE_X, non-numeric
                lib_part1 = " ".join(
                    w['text'] for w in ws
                    if LIBELLE_X - 5 <= w['x0'] < NO_LIGNE_X and not NUM_RE.match(w['text'])
                )

                # Libellé part 2: next y-row, only words at x~79 (libellé continuation)
                lib_part2 = ""
                if idx + 1 < len(y_list):
                    next_y = y_list[idx + 1]
                    next_ws = rows[next_y]
                    # Continuation row: no EAN-like word on the left, has text at libellé x
                    if next_ws and next_ws[0]['x0'] > EAN_X_MAX - 5:
                        lib_part2 = " ".join(
                            w['text'] for w in next_ws
                            if LIBELLE_X - 5 <= w['x0'] < NO_LIGNE_X
                        )

                libelle = _normalize_libelle((lib_part1 + " " + lib_part2).strip())

                # Quant en UVC: numeric token at x >= QTY_UVC_X (column ~324)
                uvc_candidates = [
                    w['text'] for w in ws
                    if w['x0'] >= QTY_UVC_X and NUM_RE.match(w['text'])
                ]
                if not uvc_candidates:
                    continue
                qty = float(uvc_candidates[0])

                if ean not in data:
                    data[ean] = {"libelle": libelle}
                data[ean][magasin] = data[ean].get(magasin, 0) + qty

    titre = "BON DE COMMANDE — MEDIDIS / MARJANE HOLDING"
    if date_cmd:
        titre += f" — {date_cmd}"
    return data, date_cmd, titre


# ─────────────────────────────────────────────
# PARSEUR MARJANE (classique)
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
                qty = float(nums[-1])

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
    for col, label in [(EAN_col, "EAN Article"), (LIB_col, "Libelle Article")]:
        c = ws.cell(2, col, label)
        c.font = _font(bold=True, color=HEADER_FG)
        c.fill = _fill(HEADER_BG)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
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
    if fmt == "medidis":
        data, date_cmd, titre = parse_medidis(pdf_path)
    elif fmt == "marjane":
        data, date_cmd, titre = parse_marjane(pdf_path)
    else:
        data, date_cmd, titre = parse_lv(pdf_path)
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