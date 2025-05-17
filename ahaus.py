import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from datetime import datetime

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
    except:
        return None, None

def get_kw(date):
    if pd.isna(date):
        return ""
    try:
        return f"KW{pd.to_datetime(date).isocalendar().week}"
    except:
        return ""

def check_zulage(comment):
    if isinstance(comment, str):
        comment_lower = comment.lower()
        return any(x in comment_lower for x in ["ahaus", "borkholzhausen", "glandorf", "alles"])
    return False

def process_file(file):
    df = pd.read_excel(file, sheet_name=0, header=None)
    df = df.iloc[4:]
    df.columns = range(df.shape[1])

    entries = []

    for _, row in df.iterrows():
        nachname_raw = row[3] if pd.notna(row[3]) else row[6]
        vorname = row[4] if pd.notna(row[4]) else row[7]
        lkw = row[11]
        datum = row[14]
        kommentar = row[15]

        if pd.notna(nachname_raw) and check_zulage(kommentar):
            monat, jahr = get_month_year(datum)
            if monat and jahr:
                nachname_check = str(nachname_raw).strip().lower()
                zulage = 0 if "zippel" in nachname_check else 20
                name = f"{nachname_raw}, {vorname}"

                ahaus_info = ""
                if isinstance(kommentar, str) and "ahaus" in kommentar.lower():
                    ahaus_info = str(row[1]) if pd.notna(row[1]) else ""

                eintrag = {
                    "Name": name,
                    "LKW": lkw,
                    "Datum": pd.to_datetime(datum, errors='coerce'),
                    "KW": get_kw(datum),
                    "Zulage": zulage,
                    "Monat": monat,
                    "Jahr": jahr,
                    "Ahaus Info": ahaus_info
                }
                entries.append(eintrag)
    return entries

def write_excel(monatsdaten):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    name_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for (monat, jahr) in sorted(monatsdaten.keys(), key=lambda x: (x[1], x[0])):
        daten = monatsdaten[(monat, jahr)]
        ws = wb.create_sheet(f"{german_months[monat]} {jahr}")
        daten.sort(key=lambda x: (x["Name"], x["Datum"]))
        current_row = 2
        current_name = None

        for eintrag in daten:
            name = eintrag["Name"]
            if name != current_name:
                if current_name is not None:
                    current_row += 1
                ws.append(["Name", "Datum", "KW", "LKW", "Zulage (€)", "Ahaus Info"])
                for col in range(1, 7):
                    cell = ws.cell(row=current_row, column=col)
                    cell.fill = header_fill
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="left")
                    cell.border = thin_border
                current_row += 1
                current_name = name

            ws.cell(row=current_row, column=1, value=name).fill = name_fill
            ws.cell(row=current_row, column=2, value=eintrag["Datum"].strftime("%d.%m.%Y"))
            ws.cell(row=current_row, column=3, value=eintrag["KW"])
            ws.cell(row=current_row, column=4, value=eintrag["LKW"])
            ws.cell(row=current_row, column=5, value=f"{eintrag['Zulage']} €")
            ws.cell(row=current_row, column=6, value=eintrag.get("Ahaus Info", ""))

            for col in range(1, 7):
                ws.cell(row=current_row, column=col).alignment = Alignment(horizontal="left")
                ws.cell(row=current_row, column=col).border = thin_border

            current_row += 1

        max_cols = 6
        for col in range(1, max_cols + 1):
            max_length = max(
                len(str(ws.cell(row=r, column=col).value)) if ws.cell(row=r, column=col).value else 0
                for r in range(1, ws.max_row + 1)
            )
            ws.column_dimensions[get_column_letter(col)].width = max_length * 2.0

        ws.row_dimensions[1].hidden = True

    wb.save(output)
    return output

# Streamlit UI
st.title("Zulage-Auswertung – 20 € für alle außer Zippel + Ahaus-Zusatzinfo")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    alle_eintraege = []
    for file in uploaded_files:
        eintraege = process_file(file)
        alle_eintraege.extend(eintraege)

    if alle_eintraege:
        monatsweise = {}
        for eintrag in alle_eintraege:
            key = (eintrag["Monat"], eintrag["Jahr"])
            monatsweise.setdefault(key, []).append(eintrag)

        excel_data = write_excel(monatsweise)
        st.download_button(
            label="📥 Excel herunterladen",
            data=excel_data.getvalue(),
            file_name="zulagen_auswertung.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Keine passenden Daten gefunden.")
