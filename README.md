# Program Review Automation: Course Enrollment & Modality Reports

This project demonstrates how I automated the production of program review reports at **Bergen Community College** to support department-level planning and scheduling. These reports were created using Python and summarize **enrollment patterns** across multiple semesters by **course, modality, and time of day**.

## Purpose

Each academic department is responsible for reviewing how its courses are scheduled and accessed. These reports help departments make informed decisions by answering key questions like:

- How many students are enrolling in each course?
- Are courses being offered online, in-person, or hybrid?
- What time of day are most students attending — morning, afternoon, or evening?

By analyzing these trends, departments can align their scheduling with student demand and optimize course delivery formats.

---

## My Contributions

- Wrote a **Python script** that loads multi-year course enrollment data and generates department-level summary tables
- Used **custom logic** to group terms by year and semester and analyze enrollment patterns
- Aggregated enrollments and section counts by:
  - **Modality** (Face-to-Face, Online, Hybrid)
  - **Time of Day** (Day vs. Evening classes)
- Automatically generated **PDF reports** for each course prefix (e.g., `ELC`, `SOC`) that include:
  - Term-by-term breakdowns
  - Grand totals
  - Program review notes and caveats
  - Pagination with dynamic headers and footers

---

## Tools Used

- **Python** – Data processing, logic, and PDF generation (`pandas`, `reportlab`)
- **Excel** – Data source for course enrollment, modality, and time-of-day flags

---

## Included Code Sample

📎 [View full script → `program_review_report_demo.py`](./program_review_report_demo.py)

> *Note: The script uses synthetic or de-identified structure. No internal data is shared in this repository.*

---

## Notes

- The actual reports are used internally at Bergen and are not public due to student/course-level data.
- However, this demo version showcases the logic and formatting structure used to build real reports.
- Program reviews help guide course offerings that align with **student access**, **enrollment trends**, and **Perkins V goals**.

---

Let me know if you'd like help generating:
- A **demo Excel file** (`demo_file_w_tod_mod.xlsx`)
- A **sample PDF output**
- A GitHub badge or thumbnail image for visual appeal
