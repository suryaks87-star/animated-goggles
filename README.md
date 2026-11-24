# School Result Analysis App

A Streamlit application for generating subject-wise result analysis from a consolidated marks Excel workbook.

## Features
- Upload a different class consolidated marks workbook whenever needed.
- Detect subjects that contain both **Internal** and **Periodic** components.
- Calculate `Internal + Periodic`, then convert the total to a percentage using the combined maximum marks.
- Show grade counts, grade percentages, performance criteria, final rating, and student-wise marks before downloading.
- Download either the selected subject report or a workbook containing all detected subjects.
- Uses the supplied result-analysis workbook as the Excel formatting template.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Create a GitHub repository, for example `school-result-analysis`.
2. Upload `app.py`, `requirements.txt`, `README.md`, and `result_analysis_template.xlsx` to the repository root.
3. Open https://share.streamlit.io/ and sign in with GitHub.
4. Choose **New app**, select the repository and branch, and set the main file path to `app.py`.
5. Deploy. The app will provide a shareable URL.

## Expected input workbook structure
The app currently expects the same consolidated-report layout as the supplied example:
- Subject headings in row 6 (merged cells across their component columns).
- Maximum marks in row 7.
- Component names such as `Internal` and `Periodic` in row 9.
- Student names in column A, starting from row 10.
- One worksheet containing the consolidated marks.

Subjects without both Internal and Periodic columns are skipped.

## Rating logic
The rating is applied sequentially, using the share of the class in each grade band:
1. **C1 and above** (C1, B2, B1, A2, A1) must be at least 75%; otherwise the result is **Needs Improvement**.
2. If the first condition passes, **B1 and above** (B1, A2, A1) is checked:
   - 50% to below 61%: **Good**
   - 61% to below 75%: **Very Good**
   - 75% or more: **Outstanding**
   - Below 50%: **Acceptable**

If the first condition passes but the B1-and-above percentage is below 50%, the final result remains **Acceptable**.

## Notes
- Blank Internal or Periodic cells are treated as zero for the calculation.
- Percentages are classified using the raw percentage, not a rounded display value.
- The included Excel template is the result-analysis format supplied for the initial setup. You can replace it with an updated template while keeping the expected layout.
