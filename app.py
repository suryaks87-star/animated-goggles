import io
import re
from copy import copy
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment

st.set_page_config(page_title="School Result Analysis", page_icon="📊", layout="wide")

TEMPLATE_PATH = Path(__file__).parent / "result_analysis_template.xlsx"
GRADE_ROWS = [
    (0, 32, "0-32", "E"),
    (33, 40, "33-40", "D"),
    (41, 50, "41-50", "C2"),
    (51, 60, "51-60", "C1"),
    (61, 70, "61-70", "B2"),
    (71, 80, "71-80", "B1"),
    (81, 90, "81-90", "A2"),
    (91, 100, "91-100", "A1"),
]
GRADE_LABELS = [g[3] for g in GRADE_ROWS]


def as_number(value):
    if value is None or value == "":
        return 0.0
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0.0


def get_subject_groups(ws):
    """Read subject blocks from the merged headings in row 6."""
    groups = []
    covered = set()
    for merged in ws.merged_cells.ranges:
        if merged.min_row == 6 and merged.max_row == 6 and merged.min_col >= 2:
            start, end = merged.min_col, merged.max_col
            name = ws.cell(6, start).value
            if name:
                groups.append((str(name).strip(), start, end))
                covered.update(range(start, end + 1))
    # Also support subjects whose heading is a single, unmerged cell.
    for col in range(2, ws.max_column + 1):
        name = ws.cell(6, col).value
        if name and col not in covered:
            groups.append((str(name).strip(), col, col))
    groups.sort(key=lambda x: x[1])
    return groups


def extract_class_info(ws):
    class_text = str(ws["A4"].value or "").replace("Class:", "").strip()
    division_text = str(ws["B4"].value or "")
    # In the provided report, Division is in D4; support either B4 or D4.
    for cell in ("B4", "D4", "C4", "E4"):
        val = ws[cell].value
        if val and "division" in str(val).lower():
            division_text = str(val).split(":", 1)[-1].strip()
            break
    if not division_text or "division" in division_text.lower():
        division_text = ""
    class_section = class_text
    if division_text:
        class_section = f"{class_text} {division_text}".strip()
    return class_text, division_text, class_section


def read_marks(uploaded_file):
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb[wb.sheetnames[0]]
    students = []
    for row in range(10, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if name is not None and str(name).strip():
            students.append((row, str(name).strip()))
    if not students:
        raise ValueError("I could not find student names in column A from row 10 onward.")

    subject_data = {}
    for subject, start, end in get_subject_groups(ws):
        internal_col = None
        periodic_col = None
        internal_max = periodic_max = 0.0
        for col in range(start, end + 1):
            component = str(ws.cell(9, col).value or "").strip().lower()
            max_mark = as_number(ws.cell(7, col).value)
            if component == "internal":
                internal_col, internal_max = col, max_mark
            elif component == "periodic":
                periodic_col, periodic_max = col, max_mark
        # Only use subjects that have both requested components.
        if internal_col is None or periodic_col is None:
            continue
        max_total = internal_max + periodic_max
        if max_total <= 0:
            continue
        records = []
        for row, name in students:
            internal_mark = as_number(ws.cell(row, internal_col).value)
            periodic_mark = as_number(ws.cell(row, periodic_col).value)
            raw_total = internal_mark + periodic_mark
            percent = (raw_total / max_total) * 100
            percent = max(0.0, min(100.0, percent))
            records.append({
                "Student Name": name,
                "Internal": internal_mark,
                "Periodic": periodic_mark,
                "Total": raw_total,
                "Maximum": max_total,
                "Percentage": percent,
                "Grade": grade_for(percent),
            })
        subject_data[subject] = {
            "records": records,
            "internal_max": internal_max,
            "periodic_max": periodic_max,
            "max_total": max_total,
        }

    if not subject_data:
        raise ValueError("No subjects with both Internal and Periodic marks were found. Please check the column headings in rows 6, 7 and 9.")
    class_text, division_text, class_section = extract_class_info(ws)
    return subject_data, students, class_text, division_text, class_section


def grade_for(percent):
    for low, high, _, grade in GRADE_ROWS:
        if low <= percent <= high:
            return grade
    return "A1" if percent > 100 else "E"


def analyse_subject(subject_info):
    records = subject_info["records"]
    strength = len(records)
    counts = {grade: sum(1 for r in records if r["Grade"] == grade) for grade in GRADE_LABELS}
    grade_percentages = {grade: (counts[grade] / strength * 100 if strength else 0) for grade in GRADE_LABELS}
    at_least_c1 = sum(grade_percentages[g] for g in ["C1", "B2", "B1", "A2", "A1"])
    at_least_b1 = sum(grade_percentages[g] for g in ["B1", "A2", "A1"])

    # Apply the thresholds in sequence. A later level is considered only
    # after the preceding level has been achieved.
    if at_least_c1 < 75:
        rating = "Needs Improvement"
    elif at_least_b1 < 50:
        rating = "Acceptable"
    elif at_least_b1 < 61:
        rating = "Good"
    elif at_least_b1 < 75:
        rating = "Very Good"
    else:
        rating = "Outstanding"

    grade_table = []
    for low, high, mark_range, grade in GRADE_ROWS:
        grade_table.append({
            "Mark Range": mark_range,
            "Grade": grade,
            "No. of students": counts[grade],
            "%": grade_percentages[grade],
        })
    return {
        "strength": strength,
        "counts": counts,
        "grade_percentages": grade_percentages,
        "at_least_c1": at_least_c1,
        "at_least_b1": at_least_b1,
        "rating": rating,
        "grade_table": grade_table,
    }


def safe_sheet_name(name):
    cleaned = re.sub(r'[\\/*?:\[\]]', "-", name).strip()[:31]
    return cleaned or "Subject"


def copy_sheet_format(wb, source_ws, new_title):
    target = wb.copy_worksheet(source_ws)
    target.title = new_title
    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target[cell.coordinate]
            if cell.has_style:
                new_cell._style = copy(cell._style)
            if cell.number_format:
                new_cell.number_format = cell.number_format
            if cell.alignment:
                new_cell.alignment = copy(cell.alignment)
            if cell.protection:
                new_cell.protection = copy(cell.protection)
    for key, dim in source_ws.column_dimensions.items():
        target.column_dimensions[key].width = dim.width
        target.column_dimensions[key].hidden = dim.hidden
    for key, dim in source_ws.row_dimensions.items():
        target.row_dimensions[key].height = dim.height
        target.row_dimensions[key].hidden = dim.hidden
    for merged in source_ws.merged_cells.ranges:
        target.merge_cells(str(merged))
    return target


def fill_analysis_sheet(ws, subject, subject_info, class_section, teacher_name=""):
    analysis = analyse_subject(subject_info)
    ws["C5"] = teacher_name
    ws["C6"] = class_section
    ws["F6"] = subject
    ws["I6"] = analysis["strength"]
    ws["H17"] = analysis["rating"]
    ws["D17"] = analysis["strength"]
    ws["E17"] = 100 if analysis["strength"] else 0

    for idx, item in enumerate(analysis["grade_table"], start=9):
        ws.cell(idx, 2).value = item["Mark Range"]
        ws.cell(idx, 3).value = item["Grade"]
        ws.cell(idx, 4).value = item["No. of students"]
        ws.cell(idx, 5).value = round(item["%"], 2)

    group_rows = [9, 11, 13, 15]
    group_values = [
        analysis["grade_percentages"]["E"] + analysis["grade_percentages"]["D"],
        analysis["grade_percentages"]["C2"] + analysis["grade_percentages"]["C1"],
        analysis["grade_percentages"]["B2"] + analysis["grade_percentages"]["B1"],
        analysis["grade_percentages"]["A2"] + analysis["grade_percentages"]["A1"],
    ]
    for row, value in zip(group_rows, group_values):
        ws.cell(row, 7).value = round(value, 2)
    ws["D17"] = analysis["strength"]
    ws["E17"] = 100 if analysis["strength"] else 0
    return analysis


def build_workbook(subject_data, class_section, selected_subject=None):
    if TEMPLATE_PATH.exists():
        wb = load_workbook(TEMPLATE_PATH)
        base = wb["SST"] if "SST" in wb.sheetnames else wb[wb.sheetnames[0]]
    else:
        raise FileNotFoundError("The result analysis template file is missing from the app folder.")

    desired = list(subject_data.keys()) if selected_subject is None else [selected_subject]
    # Retain one formatted base sheet and create a clean sheet for each subject.
    for sheet_name in list(wb.sheetnames):
        if sheet_name != base.title:
            del wb[sheet_name]
    base.title = safe_sheet_name(desired[0])
    for idx, subject in enumerate(desired):
        sheet_name = safe_sheet_name(subject)
        if idx == 0:
            ws = wb[base.title]
        else:
            ws = copy_sheet_format(wb, base, sheet_name)
        fill_analysis_sheet(ws, subject, subject_data[subject], class_section)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


st.title("School Result Analysis")
st.caption("Upload a class consolidated marks workbook, preview each subject, and download the result analysis.")

with st.sidebar:
    st.header("Upload class marks")
    uploaded = st.file_uploader("Choose consolidated marks Excel file", type=["xlsx"])
    st.caption("The marks workbook should have subject headings in row 6, maximum marks in row 7, component names in row 9, and student names in column A from row 10 onward.")

if uploaded is None:
    st.info("Upload a consolidated marks Excel file to begin.")
    st.markdown("**Calculation:** Internal marks + Periodic marks, converted to a percentage of the combined maximum marks.")
    st.stop()

try:
    subject_data, students, class_text, division_text, class_section = read_marks(uploaded)
except Exception as exc:
    st.error(f"Could not read this workbook: {exc}")
    st.stop()

st.success(f"Loaded {len(students)} students and found {len(subject_data)} subjects with Internal + Periodic marks.")
subject_names = list(subject_data.keys())
selected_subject = st.selectbox("Select subject", subject_names)
info = subject_data[selected_subject]
analysis = analyse_subject(info)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Class strength", analysis["strength"])
c2.metric("Internal maximum", info["internal_max"])
c3.metric("Periodic maximum", info["periodic_max"])
c4.metric("Combined maximum", info["max_total"])

st.subheader(f"Result Analysis — {selected_subject}")
st.caption(f"Class / Section: {class_section or 'Not identified'}")

preview = pd.DataFrame(analysis["grade_table"])
preview["%"] = preview["%"].map(lambda x: f"{x:.2f}%")
st.dataframe(preview, use_container_width=True, hide_index=True)

st.subheader("Performance criteria")
criteria = pd.DataFrame([
    {"Grade Range": "C1 and above", "Students' percentage": analysis["at_least_c1"], "Required percentage": "75% or more", "Result": "Passed" if analysis["at_least_c1"] >= 75 else "Not passed"},
    {"Grade Range": "B1 and above", "Students' percentage": analysis["at_least_b1"], "Required percentage": "50% / 61% / 75%", "Result": "Used for Good / Very Good / Outstanding"},
])
criteria["Students' percentage"] = criteria["Students' percentage"].map(lambda x: f"{x:.2f}%")
st.dataframe(criteria, use_container_width=True, hide_index=True)

rating_color = {"Needs Improvement": "red", "Acceptable": "orange", "Good": "blue", "Very Good": "green", "Outstanding": "green"}.get(analysis["rating"], "gray")
st.markdown(f"### Final rating: :{rating_color}[{analysis['rating']}]")
st.caption("Rating is applied sequentially: C1-and-above must first reach 75%; then B1-and-above is checked against 50%, 61%, and 75%.")

st.subheader("Student-wise marks preview")
st.dataframe(pd.DataFrame(info["records"]), use_container_width=True, hide_index=True)

st.divider()
st.subheader("Download result analysis")
col1, col2 = st.columns(2)
with col1:
    selected_bytes = build_workbook(subject_data, class_section, selected_subject)
    st.download_button(
        "Download selected subject Excel",
        data=selected_bytes,
        file_name=f"{safe_sheet_name(class_section)}_{safe_sheet_name(selected_subject)}_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col2:
    all_bytes = build_workbook(subject_data, class_section)
    st.download_button(
        "Download all subjects Excel",
        data=all_bytes,
        file_name=f"{safe_sheet_name(class_section)}_Complete_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
