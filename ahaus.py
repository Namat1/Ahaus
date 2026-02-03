import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

# Deutsche Monatsnamen
german_months = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

def get_month_year(date):
    if pd.isna(date):
        return None, None
    try:
        if not isinstance(date, pd.Timestamp):
            date = pd.to_datetime(date, errors="coerce")
        if pd.isna(date):
            return None, None
        return date.month, date.year
    except Exception:
        return None, None

def get_kw(date):
    if pd.isna(date):
        return ""
    try:
        return f"KW{pd.to_datetime(date).isocalendar().week}"
    except Exception:
        return ""

ZULAGE_KEYWORDS = [
    "ahaus",
    "borkholzhausen",
    "glandorf",
    "optifair",
    "opti fair",
    "edv",
    "edv fleisch",
    "elfering",
    "elfering ahaus"
]

def check_zulage(comment):
    if isinstance(comment, str):
        c = comment.lower()
        return any(k in c for k in ZULAGE_KEYWORDS)
    return False

def process_file(file):
    df = pd.read_excel(file, sheet_name=0, header=None)
    df = df.iloc[4:]  # ab Zeile 5
    df.columns = range(df.shape[1])

    entries = []

    for _, row in df.iterrows():
        # feste Spalten
        lkw = row[11] if len(row) > 11 else ""
        datum = row[14] if len(row) > 14 else pd.NaT
        kommentar = row[15] if len(row) > 15 else ""

        if not check_zulage(kommentar):
            continue

        monat, jahr = get_month_year(datum)
        if not (monat and jahr):
            continue

        # Info IMMER aus Spalte B
        info = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ""

        # 2 Fahrer-Paare: D/E und G/H
        fahrer_paare = []
        if len(row) > 4:
            fahrer_paare.append((row[3], row[4]))  # D/E
        if len(row) > 7:
            fahrer_paare.append((row[6], row[7]))  # G/H

        seen = set()

        for nachname_raw, vorname_raw in fahrer_paare:
            if pd.isna(nachname_raw) and pd.isna(vorname_raw):
                continue

            nachname = str(nachname_raw).strip() if pd.notna(nachname_raw) else ""
            vorname  = str(vorname_raw).strip() if pd.notna(vorname_raw) else ""

            if not nachname:
                continue

            name = f"{nachname}, {vorname}".strip().rstrip(",")
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)

            zulage = 0 if "zippel" in nachname.lower() else 20

            entries.append({
                "Name": name,
                "LKW": lkw,
                "Datum": pd.to_datetime(datum, errors="coerce"),
                "KW": get_kw(datum),
                "Zulage": zulage,
                "Monat": monat,
                "Jahr": jahr,
                "Info": info
            })

    return entries

def write_excel(monatsdaten):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    # Farben
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    name_fill   = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    data_fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    data_fill_light = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    total_fill  = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )
    medium_border = Border(
        left=Side(style="medium", color="1F4E78"),
        right=Side(style="medium", color="1F4E78"),
        top=Side(style="medium", color="1F4E78"),
        bottom=Side(style="medium", color="1F4E78")
    )

    def write_header(ws, r):
        headers = ["Name", "Datum", "KW", "LKW", "Zulage (€)", "Info"]
        for c, v in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = header_fill
            cell.font = Font(name="Calibri", bold=True, size=10, color="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws.row_dimensions[r].height = 20

    def write_sum(ws, r, summe):
        ws.cell(row=r, column=1, value="Gesamtzulage")
        ws.cell(row=r, column=5, value=summe)
        for c in range(1, 7):
            cell = ws.cell(row=r, column=c)
            cell.fill = total_fill
            cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = medium_border
            if c == 5:
                cell.number_format = '#,##0.00 €'
        ws.row_dimensions[r].height = 22

    for (monat, jahr) in sorted(monatsdaten.keys(), key=lambda x: (x[1], x[0])):
        daten = monatsdaten[(monat, jahr)]
        ws = wb.create_sheet(f"{german_months[monat]} {jahr}")
        daten.sort(key=lambda x: (x["Name"], x["Datum"]))

        current_row = 1
        current_name = None
        alternate_row = False
        fahrer_summe = 0

        for eintrag in daten:
            name = eintrag["Name"]

            if name != current_name:
                # vorherigen Fahrer abschließen
                if current_name is not None and fahrer_summe > 0:
                    current_row += 1
                    write_sum(ws, current_row, fahrer_summe)
                    fahrer_summe = 0
                    current_row += 2  # Summenzeile + Leerzeile

                # neuen Header schreiben
                write_header(ws, current_row, )
                current_row += 1
                current_name = name
                alternate_row = False

            fill_color = data_fill_white if alternate_row else data_fill_light

            # Name
            c1 = ws.cell(row=current_row, column=1, value=name)
            c1.fill = name_fill
            c1.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
            c1.alignment = Alignment(horizontal="left", vertical="center")
            c1.border = thin_border

            # Datum / KW / LKW
            datum_str = ""
            if pd.notna(eintrag["Datum"]):
                datum_str = eintrag["Datum"].strftime("%d.%m.%Y")

            ws.cell(row=current_row, column=2, value=datum_str)
            ws.cell(row=current_row, column=3, value=eintrag["KW"])
            ws.cell(row=current_row, column=4, value=eintrag["LKW"])

            # Zulage
            z = ws.cell(row=current_row, column=5, value=eintrag["Zulage"])
            z.number_format = '#,##0.00 €'
            z.font = Font(
                name="Calibri",
                size=10,
                color="70AD47" if eintrag["Zulage"] > 0 else "2C3E50",
                bold=True if eintrag["Zulage"] > 0 else False
            )

            # Info
            ws.cell(row=current_row, column=6, value=eintrag.get("Info", ""))

            # Styling rest
            for col in [2, 3, 4, 6]:
                cell = ws.cell(row=current_row, column=col)
                cell.fill = fill_color
                cell.font = Font(name="Calibri", size=10, color="2C3E50")
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = thin_border

            z.fill = fill_color
            z.alignment = Alignment(horizontal="right", vertical="center")
            z.border = thin_border

            ws.row_dimensions[current_row].height = 20
            current_row += 1
            alternate_row = not alternate_row
            fahrer_summe += eintrag["Zulage"]

        # letzte Summenzeile
        if current_name is not None and fahrer_summe > 0:
            write_sum(ws, current_row, fahrer_summe)

        # Spaltenbreiten
        column_min_widths = {1: 25, 2: 18, 3: 12, 4: 15, 5: 18, 6: 30}
        for col in range(1, 7):
            max_len = 0
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=col).value
                if v is None:
                    continue
                max_len = max(max_len, len(str(v)))
            width = min(max(max_len + 4, column_min_widths.get(col, 12)), 70)
            ws.column_dimensions[get_column_letter(col)].width = width

        ws.freeze_panes = "A2"

    wb.save(output)
    return output

# Streamlit UI
st.title("Zulage Ahaus")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    alle_eintraege = []
    for file in uploaded_files:
        alle_eintraege.extend(process_file(file))

    if alle_eintraege:
        monatsweise = {}
        for e in alle_eintraege:
            key = (e["Monat"], e["Jahr"])
            monatsweise.setdefault(key, []).append(e)

        excel_data = write_excel(monatsweise)
        st.download_button(
            label="📥 Excel herunterladen",
            data=excel_data.getvalue(),
            file_name="Ahaus_Auswertung.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Keine passenden Daten gefunden.")
