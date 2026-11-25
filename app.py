import io
import re
import pandas as pd
import streamlit as st

from pathlib import Path
from openpyxl import load_workbook

st.set_page_config(
    page_title="School Result Analysis",
    page_icon="📊",
    layout="wide"
)

TEMPLATE = Path(__file__).parent / "result_analysis_template.xlsx"

# Grade ranges
GRADES = [
    ("0-32", "E", 0, 32),
    ("33-40", "D", 33, 40),
    ("41-50", "C2", 41, 50),
    ("51-60", "C1", 51, 60),
    ("61-70", "B2", 61, 70),
    ("71-80", "B1", 71, 80),
    ("81-90", "A2", 81, 90),
    ("91-100", "A1", 91, 100),
]

COMPONENTS = ["Internal", "Periodic", "Term1"]


def clean_name(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def get_component(value):
    name = clean_name(value)

    aliases = {
        "internal": "Internal",
        "periodic": "Periodic",
        "term1": "Term1",
        "term": "Term1",
        "termone": "Term1"
    }

    return aliases.get(name)


def get_number(value):
    try:
        if value is None or str(value).strip().lower() in [
            "", "null", "none", "na", "n/a", "-"
        ]:
            return 0.0

        return float(str(value).replace(",", "").strip())

    except (ValueError, TypeError):
        return 0.0


def find_subjects(ws):

    subjects = []
    used_columns = set()

    for merged in ws.merged_cells.ranges:

        if merged.min_row == 6 and merged.max_row == 6:

            subject = ws.cell(6, merged.min_col).value

            if subject:
                subjects.append((
                    str(subject).strip(),
                    merged.min_col,
                    merged.max_col
                ))

                used_columns.update(
                    range(merged.min_col, merged.max_col + 1)
                )

    for col in range(2, ws.max_column + 1):

        subject = ws.cell(6, col).value

        if subject and col not in used_columns:

            subjects.append((
                str(subject).strip(),
                col,
                col
            ))

    return sorted(subjects, key=lambda x: x[1])


def read_excel(uploaded_file):

    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb[wb.sheetnames[0]]

    students = []

    for row in range(10, ws.max_row + 1):

        name = ws.cell(row, 1).value

        if name and str(name).strip():
            students.append((row, str(name).strip()))

    if not students:
        raise ValueError("Student names not found in column A.")

    subject_data = {}

    for subject, start_col, end_col in find_subjects(ws):

        subject_components = []

        for col in range(start_col, end_col + 1):

            component = get_component(ws.cell(9, col).value)
            maximum = get_number(ws.cell(7, col).value)

            if component in COMPONENTS and maximum > 0:

                subject_components.append({
                    "name": component,
                    "column": col,
                    "maximum": maximum
                })

        if not subject_components:
            continue

        subject_components.sort(
            key=lambda x: COMPONENTS.index(x["name"])
        )

        maximum_total = sum(
            item["maximum"] for item in subject_components
        )

        records = []

        for excel_row, student_name in students:

            obtained = {}
            total = 0

            for component in subject_components:

                mark = get_number(
                    ws.cell(excel_row, component["column"]).value
                )

                obtained[component["name"]] = mark
                total += mark

            percentage = (
                total / maximum_total * 100
                if maximum_total else 0
            )

            percentage = max(0, min(100, percentage))

            grade = "E"

            for _, grade_name, low, high in GRADES:

                if low <= percentage <= high:
                    grade = grade_name
                    break

            records.append({
                "Student Name": student_name,
                **obtained,
                "Total Obtained": total,
                "Maximum Marks": maximum_total,
                "Percentage": percentage,
                "Grade": grade
            })

        subject_data[subject] = {
            "records": records,
            "components": subject_components,
            "maximum": maximum_total
        }

    if not subject_data:
        raise ValueError("No valid subject marks found.")

    class_name = str(ws["A4"].value or "").replace(
        "Class:", ""
    ).strip()

    division = str(ws["B4"].value or "").replace(
        "Division:", ""
    ).strip()

    class_section = f"{class_name} {division}".strip()

    return subject_data, students, class_section


def calculate_analysis(data):

    records = data["records"]
    strength = len(records)

    grade_count = {}

    for _, grade, _, _ in GRADES:

        grade_count[grade] = sum(
            1 for student in records
            if student["Grade"] == grade
        )

    grade_percentage = {
        grade: (
            grade_count[grade] / strength * 100
            if strength else 0
        )
        for grade in grade_count
    }

    # FIRST CONDITION:
    # C1 and above must be 75% or more.

    c1_above = sum(
        grade_percentage[g]
        for g in ["C1", "B2", "B1", "A2", "A1"]
    )

    # SECOND CONDITION:
    # B1 and above percentage.

    b1_above = sum(
        grade_percentage[g]
        for g in ["B1", "A2", "A1"]
    )

    if c1_above < 75:
        rating = "Not Acceptable"

    elif b1_above < 50:
        rating = "Acceptable"

    elif b1_above < 61:
        rating = "Good"

    elif b1_above < 75:
        rating = "Very Good"

    else:
        rating = "Outstanding"

    return {
        "strength": strength,
        "grade_count": grade_count,
        "grade_percentage": grade_percentage,
        "c1_above": c1_above,
        "b1_above": b1_above,
        "rating": rating
    }


def safe_sheet_name(name):
    return re.sub(r'[\\/*?:\[\]]', "-", str(name))[:31]


# --------------------------------------------------
# CREATE RESULT ANALYSIS EXCEL
# --------------------------------------------------

def create_result_excel(subject_data, class_section, selected=None):

    if not TEMPLATE.exists():
        raise FileNotFoundError(
            "Keep result_analysis_template.xlsx in the same folder as app.py"
        )

    wb = load_workbook(TEMPLATE)

    aliases = {
        "islamiceducation": "islamic",
        "sparabic": "sparabic"
    }

    def match_sheet(subject):

        key = clean_name(subject)
        key = aliases.get(key, key)

        for sheet in wb.sheetnames:

            sheet_key = clean_name(sheet)
            sheet_key = aliases.get(sheet_key, sheet_key)

            if sheet_key == key:
                return sheet

        return None

    selected_subjects = (
        [selected] if selected else list(subject_data.keys())
    )

    sheet_mapping = {}

    for subject in selected_subjects:

        matched = match_sheet(subject)

        if matched is None:
            matched = "SST"

        sheet_mapping[subject] = matched

    required_sheets = set(sheet_mapping.values())

    # Remove unused sheets.
    for sheet in list(wb.sheetnames):

        if sheet not in required_sheets:
            del wb[sheet]

    for subject in selected_subjects:

        sheet_name = sheet_mapping[subject]
        ws = wb[sheet_name]

        analysis = calculate_analysis(subject_data[subject])

        # Header information
        ws["C6"] = class_section
        ws["F6"] = subject
        ws["I6"] = analysis["strength"]

        # Grade-wise analysis
        for row, (mark_range, grade, _, _) in enumerate(
            GRADES, start=9
        ):

            ws.cell(row, 2).value = mark_range
            ws.cell(row, 3).value = grade

            ws.cell(row, 4).value = (
                analysis["grade_count"][grade]
            )

            ws.cell(row, 5).value = round(
                analysis["grade_percentage"][grade], 2
            )

        # Performance distribution
        distributions = [
            (
                analysis["grade_percentage"]["E"]
                + analysis["grade_percentage"]["D"]
                + analysis["grade_percentage"]["C2"]
            ),

            analysis["grade_percentage"]["C1"],

            analysis["grade_percentage"]["B2"],

            (
                analysis["grade_percentage"]["B1"]
                + analysis["grade_percentage"]["A2"]
                + analysis["grade_percentage"]["A1"]
            )
        ]

        for row, value in zip(
            [9, 11, 13, 15],
            distributions
        ):

            ws.cell(row, 7).value = round(value, 2)

        # Total
        ws["B17"] = "Total"
        ws["D17"] = analysis["strength"]
        ws["E17"] = 100
        ws["H17"] = analysis["rating"]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# --------------------------------------------------
# STREAMLIT APPLICATION
# --------------------------------------------------

st.title("📊 School Result Analysis")

st.write(
    "Upload the consolidated marks Excel file, "
    "select a subject and view the complete result analysis."
)

st.sidebar.header("Upload Class Marks")

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

if uploaded_file is None:

    st.info("Please upload the consolidated marks Excel file.")
    st.stop()

try:

    subjects, students, class_section = read_excel(
        uploaded_file
    )

except Exception as error:

    st.error(f"Error reading Excel file: {error}")
    st.stop()


st.success(
    f"Successfully loaded {len(students)} students "
    f"and {len(subjects)} subjects."
)

st.write(f"**Class:** {class_section}")

# Subject selection
selected_subject = st.selectbox(
    "Select Subject",
    list(subjects.keys())
)

data = subjects[selected_subject]

analysis = calculate_analysis(data)

st.divider()

st.header(f"Result Analysis - {selected_subject}")

# Component details
cols = st.columns(len(data["components"]) + 2)

cols[0].metric(
    "Class Strength",
    analysis["strength"]
)

for index, component in enumerate(data["components"], start=1):

    cols[index].metric(
        component["name"],
        component["maximum"]
    )

cols[-1].metric(
    "Combined Maximum",
    data["maximum"]
)

st.caption(
    "Total = Internal + Periodic + Term1 (where available)"
)

# Grade analysis table
st.subheader("Grade-wise Result Analysis")

table = []

for mark_range, grade, _, _ in GRADES:

    table.append({
        "Mark Range": mark_range,
        "Grade": grade,
        "No. of Students": analysis["grade_count"][grade],
        "Percentage": round(
            analysis["grade_percentage"][grade], 2
        )
    })

st.dataframe(
    pd.DataFrame(table),
    use_container_width=True,
    hide_index=True
)

# Performance criteria
st.subheader("Performance Criteria")

criteria = pd.DataFrame([
    {
        "Criteria": "C1 and above",
        "Percentage": round(analysis["c1_above"], 2),
        "Requirement": "75% or more",
        "Status": (
            "Met" if analysis["c1_above"] >= 75
            else "Not Met"
        )
    },
    {
        "Criteria": "B1 and above",
        "Percentage": round(analysis["b1_above"], 2),
        "Requirement": (
            "50-60% Good | 61-74% Very Good | 75%+ Outstanding"
        ),
        "Status": (
            "Applied only if first condition is met"
        )
    }
])

st.dataframe(
    criteria,
    use_container_width=True,
    hide_index=True
)

# Final rating
st.subheader("Final Result")

st.markdown(
    f"## {analysis['rating']}"
)

# Student-wise marks
st.subheader("Student-wise Marks Preview")

student_df = pd.DataFrame(data["records"])

student_df["Percentage"] = student_df[
    "Percentage"
].round(2)

st.dataframe(
    student_df,
    use_container_width=True,
    hide_index=True
)

# Download
st.divider()

st.header("Download Result Analysis")

col1, col2 = st.columns(2)

with col1:

    selected_excel = create_result_excel(
        subjects,
        class_section,
        selected_subject
    )

    st.download_button(
        "Download Selected Subject",
        data=selected_excel,
        file_name=f"{selected_subject}_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col2:

    all_excel = create_result_excel(
        subjects,
        class_section
    )

    st.download_button(
        "Download All Subjects",
        data=all_excel,
        file_name="Complete_Result_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
