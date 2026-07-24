# 🏥 Hospital Emergency Room Analysis Dashboard | Excel Project

An end-to-end Excel dashboard project that analyzes Hospital Emergency Room operations — patient admissions, wait times, satisfaction scores, and department referrals — to help hospital stakeholders make faster, data-driven decisions.

![Final Dashboard](https://github.com/Shivamkr-ggv/hospital-emergency-room-dashboard-excel/blob/main/final_dashboard.png)

---

## 📌 Table of Contents

- [Project Overview](#-project-overview)
- [Business Problem](#-business-problem)
- [Dataset](#-dataset)
- [Tools & Skills Used](#-tools--skills-used)
- [Project Workflow](#-project-workflow)
  - [1. Data Cleaning & Preparation](#1-data-cleaning--preparation)
  - [2. Data Modeling (Power Pivot)](#2-data-modeling-power-pivot)
  - [3. DAX Calculated Columns](#3-dax-calculated-columns)
  - [4. Pivot Table Analysis](#4-pivot-table-analysis)
  - [5. Dashboard Design](#5-dashboard-design)
- [Key KPIs](#-key-kpis)
- [Dashboard Insights](#-dashboard-insights)
- [Repository Structure](#-repository-structure)
- [How to Use This Project](#-how-to-use-this-project)
- [Key Learnings](#-key-learnings)
- [Connect With Me](#-connect-with-me)

---

## 📖 Project Overview

Hospitals generate large volumes of daily patient data, but without proper analysis this data rarely translates into actionable insight. This project builds a **fully interactive Excel dashboard** — powered by Power Pivot, DAX, and PivotTables — that turns raw emergency room records into a single-screen decision-making tool for hospital management.

The final deliverable is a dynamic **Hospital Emergency Room Dashboard** with year-based slicers, monthly filters, and visual KPIs covering patient volume, wait time, satisfaction, and admission trends.

---

## 🎯 Business Problem

> **We need to create a Hospital Emergency Room Analysis Dashboard to improve efficiency and provide useful insights. This dashboard will help stakeholders monitor, analyze, and make better decisions for managing patients and improving services.**

The dashboard needed to answer the following business questions:

- How many patients were **admitted vs. not admitted**?
- What is the **age distribution** of patients visiting the ER?
- What percentage of patients were seen **within 30 minutes** (timeliness)?
- What is the **gender-wise split** of ER patients?
- Which **departments** receive the highest number of referrals?
- How do patient volume, wait time, and satisfaction trend **month-by-month**?

---

## 🗃 Dataset

**File:** `Hospital_Emergency_Room_Data.csv`
**Records:** ~9,200 patient visits (2023–2024)

| Column | Description |
|---|---|
| Patient Id | Unique identifier for each patient visit |
| Patient Admission Date | Date the patient was admitted to the ER |
| Patient First Initial / Last Name | Patient identity fields |
| Patient Gender | Male / Female |
| Patient Age | Age of the patient |
| Patient Race | Patient demographic race |
| Department Referral | Department the patient was referred to (Cardiology, Orthopedics, etc.) |
| Patient Admission Flag | Whether the patient was admitted (TRUE/FALSE) |
| Patient Satisfaction Score | Score out of 10 given by the patient |
| Patient Waittime | Wait time in minutes before being attended to |

*Note: Raw data required cleaning (missing satisfaction scores, inconsistent date formats) before modeling.*

---

## 🛠 Tools & Skills Used

- **Microsoft Excel** (Power Pivot, Power Query)
- **DAX (Data Analysis Expressions)** for calculated columns
- **PivotTables & PivotCharts**
- **Data Modeling** (Star Schema — Fact + Calendar dimension table)
- **Slicers & Interactive Filters**
- **Dashboard Design Principles** (KPI cards, chart layout, color theory)

---

## 🔄 Project Workflow

### 1. Data Cleaning & Preparation
The raw CSV file was imported into Excel and loaded into the Data Model. Fields were checked for blanks, duplicate records, and inconsistent formats (e.g., wait time and satisfaction score nulls) before being used in calculations.

### 2. Data Modeling (Power Pivot)
A **Calendar Table** was created using Power Query's `List.Dates` function to enable proper time-intelligence and month/year-based slicing:

```
= List.Dates(#date(2023,01,01), 731, #duration(1,0,0,0))
```

The Hospital Emergency Room dataset was then related to the `Calendar_Table` on the **Date** field, forming a clean one-to-many relationship (Calendar → Fact Table) inside Power Pivot's Data Model.

![Table Relationships](https://github.com/Shivamkr-ggv/hospital-emergency-room-dashboard-excel/blob/main/table_relationships.png)

The full data model, including all fields loaded into Power Pivot, is shown below:

![Power Pivot Data Model](https://github.com/Shivamkr-ggv/hospital-emergency-room-dashboard-excel/blob/main/power_pivot_data_model.png)

### 3. DAX Calculated Columns
Two custom calculated columns were created using DAX to enable grouped analysis:

**Patient Age Group:**
```dax
= IF([Patient Age] >= 70, "70-79",
  IF([Patient Age] >= 60, "60-69",
  IF([Patient Age] >= 45, "45-59",
  IF([Patient Age] >= 30, "30-44",
  IF([Patient Age] >= 15, "15-29",
  IF([Patient Age] >= 5, "05-14", "0-4"))))))
```

**Patient Attend Status (Timeliness Flag):**
```dax
= IF([Patient Waittime] < 30, "Within Time", "Delay")
```

These fields power the **Age Group distribution chart** and the **On-Time vs. Delay** pie chart in the final dashboard.

### 4. Pivot Table Analysis
Multiple PivotTables were built on top of the data model to summarize:
- Total patients, average wait time, and average satisfaction score
- Admission status (Admitted vs. Not Admitted) with % share
- Patient count by gender, age group, and department referral
- Day-wise trend of patient volume, average wait time, and satisfaction score across the year

![Pivot Table Report](https://github.com/Shivamkr-ggv/hospital-emergency-room-dashboard-excel/blob/main/pivot_table.png)

### 5. Dashboard Design
All pivot outputs were converted into charts (bar, pie, sparklines) and assembled onto a single dashboard sheet with a **Year slicer (2023/2024)** and **Month buttons (Jan–Dec)** for interactivity, giving stakeholders a self-service reporting tool.

---

##  Key KPIs

| KPI | Value |
|---|---|
| **Total Number of Patients** | 9,216 |
| **Average Wait Time** | 35.26 minutes |
| **Average Patient Satisfaction Score** | 4.99 / 10 |
| **Admitted Patients** | 4,612 (50.04%) |
| **Not Admitted Patients** | 4,604 (49.96%) |
| **On-Time (< 30 min) Patients** | 59% |
| **Delayed (≥ 30 min) Patients** | 41% |

---

##  Dashboard Insights

- **Admissions are almost evenly split** — 50.04% of ER patients were admitted vs. 49.96% not admitted, indicating the ER handles a balanced mix of critical and non-critical cases.
- **41% of patients experienced a delay** (waited 30+ minutes), highlighting an area for operational improvement in patient triage/staffing.
- **Gender distribution is nearly balanced** (51% Female, 49% Male), so ER demand isn't gender-skewed.
- **Age group 20–29 is the most frequent ER visitor group** (1,207 patients), followed closely by 30–39 and 60–69 — useful for resource planning by age-specific care needs.
- **The majority of referrals (5,000+) have no specific department ("None")**, i.e., handled directly in general ER care, while **General Practice** and **Orthopedics** are the top specialist referral departments.
- **October had the lowest patient volume (1,048)** while **March had the highest (1,207)**, useful for identifying seasonal staffing needs.

---

##  Repository Structure

```
Hospital-ER-Dashboard-Excel-Project/
│
├── README.md                                         # Project documentation (this file)
├── Hospital_Emergency_Room_Analytics.xlsx            # Final Excel file (Data, Model, Pivots, Dashboard)
├── Hospital_Emergency_Room_Data.csv                  # Raw dataset
├── Hospital-ER-Dashboard_Problem-Statement.pdf       # Business requirement / problem statement deck
│
└── screenshots/
    ├── final_dashboard.png                  # Final interactive dashboard
    ├── pivot_table.png                      # PivotTable report view
    ├── power_pivot_data_model.png           # Power Pivot data model view
    └── table_relationships.png              # Fact-to-Calendar table relationship
```

---

##  How to Use This Project

1. Clone or download this repository.
2. Open `Hospital_Emergency_Room_Analytics.xlsx` in **Microsoft Excel (2016 or later, with Power Pivot enabled)**.
3. Go to the **Dashboard** sheet to interact with the Year slicer and Month filters.
4. Explore the **Pivot_Table_Report** and **Daily Patients Analysis** sheets to see the underlying PivotTable logic.
5. Check the **Data** tab in Power Pivot (`Power Pivot → Manage`) to review the data model and DAX formulas.

---

##  Key Learnings

- Building a **star-schema data model** in Excel using Power Pivot instead of relying only on flat VLOOKUPs.
- Writing **DAX calculated columns** for dynamic grouping (age bands, timeliness flags).
- Designing a **recruiter/stakeholder-friendly dashboard layout** — KPI cards up top, filters accessible, charts grouped logically.
- Translating a **business problem statement** directly into measurable KPIs and visuals.

---

##  Connect With Me

If you found this project useful or have feedback, feel free to connect:

- **LinkedIn:** www.linkedin.com/in/shivamkumar-data
- **Email:** shivaraghav234@gmail.com

⭐ If you like this project, consider giving this repository a star!
