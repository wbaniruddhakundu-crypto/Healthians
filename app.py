import re
from io import BytesIO
from collections import defaultdict

import fitz
import openpyxl
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inter-Run Precision Populator", layout="wide")
st.title("Healthians – Inter-Run Precision Data Populator")
st.write("Upload the existing Excel precision template and the five daily PDF assay reports. The app keeps the template dates and formatting, and fills up to five result replicates per analyte per day.")

# Result units commonly present in the assay reports. Extend this list if a new unit appears.
UNIT_PATTERN = r"(?:ng/mL|pg/mL|ug/dL|µg/dL|mg/dL|g/dL|U/mL|mIU/mL|uIU/mL|µIU/mL|mIU/L|IU/L|pmol/L|nmol/L|mmol/L|µmol/L|umol/L|%)"
RESULT_RE = re.compile(rf"^\s*(-?\d+(?:,\d{{3}})*(?:\.\d+)?)\s*({UNIT_PATTERN})\s*$", re.I)


def day_from_name(name):
    m = re.search(r"day\s*[-_ ]?([1-5])", name, re.I)
    return int(m.group(1)) if m else None


def parse_pdf(pdf_bytes):
    """Extract analyte sections and numeric result/unit lines from a text-based assay PDF."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for page in doc:
        lines.extend(page.get_text("text").splitlines())
    doc.close()

    data = defaultdict(list)
    current = None
    for i, raw in enumerate(lines):
        line = raw.strip()
        # In this report layout, an analyte title is immediately followed by
        # the column headings Patient Name / SID / Result.
        lookahead = " ".join(x.strip().lower() for x in lines[i+1:i+8])
        if line and len(line) <= 24 and not RESULT_RE.match(line):
            cleaned = re.sub(r"[^A-Za-z0-9_+.-]", "", line)
            if ("patient name" in lookahead and "sid" in lookahead and "result" in lookahead
                    and cleaned.upper() not in {"ASSAYREPORT", "ROUTINE", "TUMOR"}):
                current = cleaned
        match = RESULT_RE.match(line)
        if current and match:
            value = float(match.group(1).replace(",", ""))
            unit = match.group(2)
            data[current].append((value, unit))
    return data


def safe_sheet_name(name, used):
    name = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Analyte"
    base, candidate, n = name, name, 1
    while candidate in used:
        suffix = f"_{n}"
        candidate = base[:31-len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def populate(template_bytes, day_files):
    wb0 = openpyxl.load_workbook(BytesIO(template_bytes))
    template = wb0[wb0.sheetnames[0]]

    # Read all daily PDFs; the filename is used to identify Day 1–Day 5.
    by_day = {}
    parse_notes = []
    for filename, content in day_files:
        day = day_from_name(filename)
        if day is None:
            parse_notes.append(f"Could not identify day from filename: {filename}")
            continue
        if day in by_day:
            parse_notes.append(f"More than one PDF identified as Day {day}; using the first file: {by_day[day][0]}")
            continue
        by_day[day] = (filename, parse_pdf(content))

    analytes = set()
    for _, parsed in by_day.values():
        analytes.update(parsed.keys())
    if not analytes:
        raise ValueError("No assay result sections were detected. Check that the PDFs contain selectable text (not scanned images).")

    # Preserve the template sheet as the first analyte sheet and clone it for the rest.
    first = True
    used = set()
    for analyte in sorted(analytes, key=str.upper):
        if first:
            ws = template
            first = False
        else:
            ws = wb0.copy_worksheet(template)
        ws.title = safe_sheet_name(analyte, used)

        # Fill analyte labels without changing the date cells or the overall layout.
        ws["C4"] = analyte
        ws["C6"] = analyte

        unit = ""
        for _, parsed in by_day.values():
            vals = parsed.get(analyte, [])
            if vals:
                unit = vals[0][1]
                break
        if unit:
            ws["D6"] = f"Unit: ({unit})"

        # The existing template has five day columns B:F and five replicate rows 11:15.
        for day in range(1, 6):
            col = day + 1
            vals = by_day.get(day, (None, {}))[1].get(analyte, []) if day in by_day else []
            for rep in range(5):
                cell = ws.cell(row=11 + rep, column=col)
                cell.value = vals[rep][0] if rep < len(vals) else None

        # Keep template styling; standardize formula ranges to the five replicate rows.
        for col in range(2, 7):
            letter = get_column_letter(col)
            ws.cell(17, col).value = f'=COUNT({letter}11:{letter}15)'
            ws.cell(18, col).value = f'=IFERROR(AVERAGE({letter}11:{letter}15),"")'
            ws.cell(19, col).value = f'=IFERROR(STDEV({letter}11:{letter}15),"")'
            ws.cell(20, col).value = f'=IFERROR({letter}19/{letter}18,"")'
        # Inter-run summary formulas (all available observations across five days).
        ws["E25"] = "=COUNT(B11:F15)"
        ws["E26"] = '=IFERROR(AVERAGE(B11:F15),"")'
        ws["E27"] = '=IFERROR(STDEV(B11:F15),"")'
        ws["E28"] = '=IFERROR(E27/E26,"")'

    output = BytesIO()
    wb0.calculation.fullCalcOnLoad = True
    wb0.calculation.forceFullCalc = True
    wb0.calculation.calcMode = "auto"
    wb0.save(output)
    return output.getvalue(), sorted(analytes, key=str.upper), parse_notes, sorted(by_day)


template_file = st.file_uploader("1. Upload your Excel precision template (.xlsx)", type=["xlsx"])
pdf_files = st.file_uploader("2. Upload Day 1–Day 5 PDF reports (filenames should contain Day 1, Day 2, etc.)", type=["pdf"], accept_multiple_files=True)

if st.button("Generate completed precision workbook", type="primary"):
    if not template_file or not pdf_files:
        st.error("Please upload the Excel template and at least one daily PDF.")
    else:
        try:
            payload = [(f.name, f.getvalue()) for f in pdf_files]
            result, analytes, notes, days = populate(template_file.getvalue(), payload)
            st.success(f"Workbook created: {len(analytes)} analyte sheets; reports detected for Day(s): {', '.join(map(str, days))}.")
            st.download_button("Download completed Excel workbook", data=result,
                               file_name="Healthians_Inter_Run_Precision_Daywise.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.write("**Analytes detected:** " + ", ".join(analytes))
            if notes:
                st.warning("\n".join(notes))
            st.info("Only the first five result values in each analyte section are used per day. Missing days/results remain blank. The date cells in the uploaded template are not changed. Please review the output against the source reports before validation.")
        except Exception as exc:
            st.exception(exc)
