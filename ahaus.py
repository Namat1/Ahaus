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
            date = pd.to_datetime(date, errors='coerce')
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
        comment_lower = comment.lower()
        return any(k in comment_lower for k in ZULAGE_KEYWORDS)
    return False

def process_file(file):
    df = pd.read_excel(file, sheet_name=0, header=None)
    df = df.iloc[4:]  # ab Zeile 5
    df.columns = range(df.shape[1])

    entries = []

    for _, row in df.iterrows():
        lkw = row[11]   # Spalte L
        datum = row[14] # Spalte O
        kommentar = row[15]  # Spalte P

        # nur wenn Kommentar zulagenrelevant ist
        if not check_zulage(kommentar):
            continue

        monat, jahr = get_month_year(datum)
        if not (monat and jahr):
            continue

        # Info soll IMMER aus Spalte B kommen
        info = str(row[1]) if pd.notna(row[1]) else ""

        # ggf. 2 Fahrer in einer Zeile: (D/E) und (G/H)
        fahrer_paare = [
            (row[3], row[4]),  # Nachname D, Vorname E
            (row[6], row[7])   # Nachname G, Vorname H
        ]

        seen_names = set()  # verhindert doppelte Einträge, falls beide Paare identisch sind

        for nachname_raw, vorname_raw in fahrer_paare:
            if pd.isna(nachname_raw) and pd.isna(vorname_raw):
                continue

            nachname = str(nachname_raw).strip() if pd.notna(nachname_raw) else ""
            vorname = str(vorname_raw).strip() if pd.notna(vorname_raw) else ""

            if not nachname:
                continue

            name = f"{nachname}, {vorname}".strip().rstrip(",")
            name_key = name.lower()
            if name_key in seen_names:
                continue
            seen_names.add(name_key)

            # Zulage pro Fahrer
            zulage = 0 if "zippel" in nachname.lower() else 20

            entries.append({
                "Name": name,
                "LKW": lkw,
                "Datum": pd.to_datetime(datum, errors='coerce'),
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

    # Moderne Farbpalette
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")  # Hellblau
    name_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")    # Mittelblau
    data_fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    data_fill_light = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    total_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    # Rahmen
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    medium_border = Border(
        left=Side(style='medium', color='1F4E78'),
        right=Side(style='medium', color='1F4E78'),
        top=Side(style='medium', color='1F4E78'),
        bottom=Side(style='medium', color='1F4E78')
    )

    for (monat, jahr) in sorted(monatsdaten.keys(), key=lambda x: (x[1], x[0])):
        daten = monatsdaten[(monat, jahr)]
        ws = wb.create_sheet(f"{german_months[monat]} {jahr}")
        daten.sort(key=lambda x: (x["Name"], x["Datum"]))

        current_row = 2
        current_name = None
        alternate_row = False
        fahrer_summe = 0

        for eintrag in daten:
            name = eintrag["Name"]

            # neuer Fahrer -> vorherige Summenzeile schreiben
            if name != current_name:
                if current_name is not None and fahrer_summe > 0:
                    ws.cell(row=current_row, column=1, value="Gesamtzulage")
                    ws.cell(row=current_row, column=5, value=fahrer_summe)

                    for col in range(1, 7):
                        cell = ws.cell(row=current_row, column=col)
                        cell.fill = total_fill
                        cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.border = medium_border
                        if col == 5:
                            cell.number_format = '#,##0.00 €'

                    ws.row_dimensions[current_row].height = 22
                    current_row += 1
                    fahrer_summe = 0

                if current_name is not None:
                    current_row += 1  # Leerzeile zwischen Personen

                # Header
                ws.append(["Name", "Datum", "KW", "LKW", "Zulage (€)", "Info"])
                for col in range(1, 7):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = header_fill
                    cell.font = Font(name="Calibri", bold=True, size=10, color="1F4E78")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border

                ws.row_dimensions[current_row].height = 20
                current_row += 1
                current_name = name
                alternate_row = False

            # Datenzeile
            fill_color = data_fill_white if alternate_row else data_fill_light

            # Name-Spalte
            name_cell = ws.cell(row=current_row, column=1, value=name)
            name_cell.fill = name_fill
            name_cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            name_cell.border = thin_border

            # Restliche Zellen
            ws.cell(row=current_row, column=2, value=eintrag["Datum"].strftime("%d.%m.%Y") if pd.notna(eintrag["Datum"]) else "")
            ws.cell(row=current_row, column=3, value=eintrag["KW"])
            ws.cell(row=current_row, column=4, value=eintrag["LKW"])

            zulage_cell = ws.cell(row=current_row, column=5, value=eintrag["Zulage"])
            zulage_cell.number_format = '#,##0.00 €'
            zulage_cell.font = Font(
                name="Calibri",
                size=10,
                color="70AD47" if eintrag["Zulage"] > 0 else "2C3E50",
                bold=True if eintrag["Zulage"] > 0 else False
            )

            ws.cell(row=current_row, column=6, value=eintrag.get("Info", ""))

            # Styling (außer Name)
            for col in [2, 3, 4, 6]:
                cell = ws.cell(row=current_row, column=col)
                cell.fill = fill_color
                cell.font = Font(name="Calibri", size=10, color="2C3E50")
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = thin_border

            # Zulage styling
            zulage_cell.fill = fill_color
            zulage_cell.alignment = Alignment(horizontal="right", vertical="center")
            zulage_cell.border = thin_border

            ws.row_dimensions[current_row].height = 20
            current_row += 1
            alternate_row = not alternate_row

            fahrer_summe += eintrag["Zulage"]

        # letzte Summenzeile
        if current_name is not None and fahrer_summe > 0:
            ws.cell(row=current_row, column=1, value="Gesamtzulage")
            ws.cell(row=current_row, column=5, value=fahrer_summe)

            for col in range(1, 7):
                cell = ws.cell(row=current_row, column=col)
                cell.fill = total_fill
                cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = medium_border
                if col == 5:
                    cell.number_format = '#,##0.00 €'

            ws.row_dimensions[current_row].height = 22

        # Spaltenbreiten
        column_min_widths = {
            1: 25,  # Name
            2: 18,  # Datum
            3: 12,  # KW
            4: 15,  # LKW
            5: 18,  # Zulage
            6: 30   # Info
        }

        max_cols = 6
        for col in range(1, max_cols + 1):
            max_length = max(
                len(str(ws.cell(row=r, column=col).value)) if ws.cell(row=r, column=col).value else 0
                for r in range(1, ws.max_row + 1)
            )
            calculated_width = max_length + 4
            min_width = column_min_widths.get(col, 12)
            adjusted_width = max(calculated_width, min_width)
            adjusted_width = min(adjusted_width, 70)
            ws.column_dimensions[get_column_letter(col)].width = adjusted_width

        # Freeze Panes
        ws.freeze_panes = "A3"

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
        for eintrag in alle_eintraege:
            key = (eintrag["Monat"], eintrag["Jahr"])
            monatsweise.setdefault(key, []).append(eintrag)

        excel_data = write_excel(monatsweise)
        st.download_button(
            label="📥 Excel herunterladen",
            data=excel_data.getvalue(),
            file_name="Ahaus_Auswertung.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Keine passenden Daten gefunden.")
