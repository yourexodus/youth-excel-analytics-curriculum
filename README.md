# 📊 Excel Data Analytics Course Syllabus (12-Day Program)

> **Target Audience:** High School Students (Ages 15–18)  
> **Pacing:** 12 Days (60–90 Minutes per Day | 2 Days per Unit)  
> **Prerequisites:** Windows 10/11 (64-bit), Microsoft Excel, Google Chrome  

---

## 🎨 Course Overview & Core Values

Welcome to **Data Analytics**! In this hands-on 12-day course, you'll discover how to turn raw, messy datasets into actionable intelligence using Microsoft Excel. Spanning 6 core units over 12 sessions, you'll master data hygiene, dynamic formulas, conditional logic, PivotTable summaries, and interactive dashboards, culminating in a professional data portfolio published on GitHub.

| 🌟 **Our Core Values** | 🎯 **What You Will Learn** | 💻 **System Requirements** |
| :--- | :--- | :--- |
| 🟢 **Respectful** & Supportive | 🔍 Connect & organize raw data sources | **OS:** Windows 10, 11, or Server 2016+ (64-bit) |
| 💡 **Curious** & Engaged | 🧹 Identify & correct real-world data issues | **CPU:** 1 GHz or faster 64-bit (x64) |
| 🛡️ **Safe for Learning** | ⚡ Transform & combine dynamic datasets | **RAM:** 2 GB min (4 GB recommended) |
| 🚀 Tech Progress & Growth | 📈 Build an executive portfolio | **Display:** 1440x900 or 1600x900 (100% scale) |

---

## 📁 Course Datasets & Interactive Tools

Students select 1 primary dataset from the local `data/` folder to analyze across all 12 sessions:

* 🎮 **Dataset 1:** [Gamer Gear & Tech Store Sales (`data/01_gamer_gear_sales_messy.csv`)](data/01_gamer_gear_sales_messy.csv)
* 🏆 **Dataset 2:** [Esports Tournament & Player Stats (`data/02_esports_player_stats.csv`)](data/02_esports_player_stats.csv)
* 👟 **Dataset 3:** [Sneaker & Streetwear Resale Tracker (`data/03_sneaker_streetwear_resale.csv`)](data/03_sneaker_streetwear_resale.csv)
* 🍿 **Dataset 4:** [Streaming Movies & Anime Analytics (`data/04_streaming_movies_shows.csv`)](data/04_streaming_movies_shows.csv)
* ⚡ **Dataset 5:** [High School Athletic Combine Performance (`data/05_high_school_fitness_combine.csv`)](data/05_high_school_fitness_combine.csv)

🛠️ **Course Tools & Web Applications:**
* 🧹 **Data Cleaning Answer Key:** [Master Data Cleaning Guide](data_cleaning_guide.md)
* 📝 **README Generator Web App:** [Build My README App](https://script.google.com/macros/s/AKfycbzVnDWxPtew1hFOfx6IiaCXZoirqNJS802LywmHoAdtsIl9WCjAwDAaVIlteh56dV2m/exec)

---

## 📅 12-Day Detailed Unit Schedule

---

### 🟢 UNIT 1: Data Architecture & Grid Foundations (Days 1 & 2)

#### 📅 DAY 1 (Unit 1, Part A): Interface Mastery & Data Type Audit (60–90 Mins)
* **Focus:** Interface Navigation, Boundaries, Data Types & View Locking
* **Lab:** Grid Navigation & Formatting Setup
```text
[EXERCISE 1.1] Grid Navigation & Boundary Mapping
1. Open your selected dataset from the 'data/' folder in Excel.
2. Select cell A1. Press CTRL + DOWN ARROW to find total rows. Record exact count.
3. Press CTRL + RIGHT ARROW to identify all attribute columns.
4. Highlight Row 1 -> View Tab -> Freeze Panes -> Click 'Freeze Top Row'.
5. Scroll down 100 rows to ensure headers remain locked.

[EXERCISE 1.2] Column Type Classification & Explicit Formatting
1. On a new tab named 'Audit_Log', create headers: 'Column_Name' and 'Data_Type'.
2. Classify each column as: Text, Currency, Decimal, Integer, or Date.
3. Format date columns explicitly to Short Date (YYYY-MM-DD).
4. Format monetary columns to Currency ($) with 2 decimals.
```

#### 📅 DAY 2 (Unit 1, Part B): Deep-Dive Data Audit (60–90 Mins)
* **Focus:** Data Quality Diagnostics & Mixed Data Type Identification
* **Lab:** Spotting Mixed Types & Anomaly Logging
```text
[EXERCISE 1.3] Data Detective Challenge: Mixed Type Diagnostics
1. Inspect numerical columns (Price, GPA, Scores) for text-stored values.
2. Use Conditional Formatting: Home -> Conditional Formatting -> Highlight Cells Rules -> Text That Contains "$".
3. Identify cells with irregular formatting (e.g., "$49.99" mixed into numeric columns).
4. Create an anomaly log on Sheet2 listing cell locations and error descriptions.
```
* **Unit 1 Deliverable:** Clean workbook with locked headers, formatted numbers/dates, and a completed Data Type Audit Log.

---

### 🔵 UNIT 2: Data Hygiene & Cleaning Workflows (Days 3 & 4)

#### 📅 DAY 3 (Unit 2, Part A): CSV Ingestion & Deduplication (60–90 Mins)
* **Focus:** Clean Data Import, Deduplication, & Range Setup
* **Lab:** Ingestion & Duplicate Scrubbing
```text
[EXERCISE 2.1] CSV Import via Power Query Engine
1. Open Excel -> Data Tab -> Get Data -> From File -> From Text/CSV.
2. Select your dataset file and set Data Type Detection to 'Based on entire dataset'. Click Load.
3. Press CTRL + A -> Data Tab -> Remove Duplicates.
4. Verify header recognition, click OK, and record how many duplicate rows were removed.

[EXERCISE 2.2] White-Space & Case Cleanup
1. Insert a blank column titled 'Clean_String'.
2. Enter formula: =PROPER(TRIM(B2)) to remove leading/trailing spaces and correct capitalization.
3. Fill down the column -> Copy -> Paste Special as Values over the original column. Delete helper column.
```

#### 📅 DAY 4 (Unit 2, Part B): Diagnostic Filtering & Standardization (60–90 Mins)
* **Focus:** AutoFilters, Categorical Standardization, & Blank Value Handling
* **Lab:** Filter Diagnostics & Anomaly Correction
```text
[EXERCISE 2.3] Data Detective Challenge: Categorical Standardization
1. Apply AutoFilters across headers (CTRL + SHIFT + L).
2. Open filter dropdowns for categorical columns (Region, Genre, Brand).
3. Search for case splits (e.g., "NORTH" vs "North") or typos.
4. Use Find & Replace (CTRL + H) to unify all regional/category terms into Proper Case.
5. Filter by 'Blanks' in numeric columns and impute default values ('0' or 'N/A').
```
* **Unit 2 Deliverable:** Deduplicated, scrubbed, and standardized dataset with zero trailing spaces or split category names.

---

### 🟡 UNIT 3: Analytics Formulas & Conditional Logic (Days 5 & 6)

#### 📅 DAY 5 (Unit 3, Part A): Statistical Metrics & Summary Sheets (60–90 Mins)
* **Focus:** Aggregation Functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`)
* **Lab:** Building the Central Summary Sheet
```text
[EXERCISE 3.1] Statistical Summary Benchmarks
Create a new tab named 'KPI_Summary' and build calculations using named ranges:
1. Total Aggregated Revenue/Score:  =SUM(CleanData!E2:E200)
2. Average Metric Value:            =AVERAGE(CleanData!E2:E200)
3. Total Valid Record Count:        =COUNT(CleanData!E2:E200)
4. Maximum Value Recorded:          =MAX(CleanData!E2:E200)
5. Minimum Value Recorded:          =MIN(CleanData!E2:E200)
```

#### 📅 DAY 6 (Unit 3, Part B): Conditional Aggregations & Lookups (60–90 Mins)
* **Focus:** Logical Functions (`IF`, `SUMIF`, `COUNTIF`) and Multi-Table Merges (`XLOOKUP`)
* **Lab:** Advanced Logic & Cross-Table Retrieval
```text
[EXERCISE 3.2] Conditional Logic & Tiering
1. Write conditional aggregations for specific categories:
   =SUMIF(CleanData!C2:C200, "Apparel", CleanData!E2:E200)
2. Calculate performance counts exceeding averages:
   =COUNTIF(CleanData!E2:E200, ">" & B2)
3. Build a calculated column 'Tier_Rating' using nested IF logic:
   =IF(E2>=100, "Elite", IF(E2>=50, "Standard", "Underperforming"))

[EXERCISE 3.3] Data Detective Challenge: Relational XLOOKUP
1. Create a reference table listing IDs and target goals on Sheet3.
2. In your main dataset, write an XLOOKUP formula to retrieve target benchmarks:
   =XLOOKUP(A2, TargetSheet!A$2:A$20, TargetSheet!B$2:B$20, "Missing Goal")
```
* **Unit 3 Deliverable:** A robust summary worksheet with statistical KPIs, conditional tiering logic, and working `XLOOKUP` merges.

---

### 🟣 UNIT 4: Summarizing Data with PivotTables (Days 7 & 8)

#### 📅 DAY 7 (Unit 4, Part A): Dynamic Tables & Pivot Fundamentals (60–90 Mins)
* **Focus:** Structured Tables (`CTRL + T`), Value Summaries, & Matrix Layouts
* **Lab:** Table Conversion & Initial Pivot Matrix
```text
[EXERCISE 4.1] Range Conversion & Basic Pivot Setup
1. Click inside your clean dataset -> Press CTRL + T -> Confirm headers -> Click OK.
2. Name table 'CleanDataTable' in Table Design tab.
3. Click Insert -> PivotTable -> Choose 'New Worksheet' (Name: 'Pivot_Analysis').
4. Drag primary Category to ROWS and main Numeric Field to VALUES (Set to SUM or AVERAGE).

[EXERCISE 4.2] Multi-Dimensional Matrix Setup
1. Drag a secondary attribute into COLUMNS to build a dynamic cross-tabulation grid.
2. Right-click any value cell -> Show Values As -> Select '% of Column Total'.
```

#### 📅 DAY 8 (Unit 4, Part B): Grouping & Interactive Slicers (60–90 Mins)
* **Focus:** Date/Number Grouping, Timeline Filters, & Interactive Slicers
* **Lab:** Advanced Pivot Modeling & Interactivity
```text
[EXERCISE 4.3] Data Detective Challenge: Slicer Dashboard Mechanics
1. Drag Date field to ROWS. Right-click a date cell -> Group -> Select 'Months' & 'Years'.
2. Click inside PivotTable -> PivotTable Analyze -> Insert Slicer.
3. Select 2 categorical attributes (e.g., Rank_Tier, Grade_Level, Region) to create visual filter buttons.
4. Add a Timeline slicer for date filtering. Format slicers into multi-column layouts.
```
* **Unit 4 Deliverable:** A dedicated `Pivot_Analysis` tab featuring multi-dimensional matrix tables, custom groupings, and interactive slicers.

---

### 🟧 UNIT 5: Data Visualization & Executive Dashboards (Days 9 & 10)

#### 📅 DAY 9 (Unit 5, Part A): Chart Selection & Formatting (60–90 Mins)
* **Focus:** Line Charts (Time Trends), Column Charts (Comparison), & Scatter Plots (Correlation)
* **Lab:** Visual Construction & Formatting
```text
[EXERCISE 5.1] Time Trends & Comparison Graphics
1. Select monthly PivotTable data -> Insert Tab -> Line Chart with Markers. Title: "Monthly Trend Analysis".
2. Add vertical axis title ("Revenue ($)").
3. Select categorical totals -> Insert Clustered Column Chart.
4. Highlight top-performing bar with an accent color; format remaining bars in gray.

[EXERCISE 5.2] Scatter Plot Correlation Testing
1. Select two numeric fields (e.g., Dash Time vs Vertical Jump, Retail vs Resale Price).
2. Click Insert -> Scatter Chart. Add clear X and Y axis titles.
3. Right-click markers -> Add Trendline -> Display R-squared value to assess correlation.
```

#### 📅 DAY 10 (Unit 5, Part B): Executive Dashboard Assembly (60–90 Mins)
* **Focus:** Dashboard Layout, De-cluttering, & Unified Slicer Connections
* **Lab:** Building the Final Dashboard Screen
```text
[EXERCISE 5.3] Data Detective Challenge: Dashboard Assembly
1. Create a sheet named 'Executive_Dashboard'. Disable gridlines (View -> Uncheck Gridlines).
2. Move Line Chart, Column Chart, Scatter Plot, and Slicers onto the dashboard tab.
3. Align visual components using Grid Snapping (Page Layout -> Align to Grid).
4. Connect Slicers to all charts: Right-click Slicer -> Report Connections -> Check all PivotTables.
```
* **Unit 5 Deliverable:** A clean, ungridlined `Executive_Dashboard` tab with 3 dynamic charts linked to interactive slicers.

---

### 🚀 UNIT 6: Portfolio Assembly & GitHub Deployment (Days 11 & 12)

#### 📅 DAY 11 (Unit 6, Part A): README Generator Web App Documentation (60–90 Mins)
* **Focus:** Workbook Auditing & Project Documentation Generation
* **Lab:** Documentation Web App Execution

```text
[EXERCISE 6.1] Interactive README Documentation
1. Launch the Readme Code Generator web app in your Chrome browser
* [Readme Code Generator](https://script.google.com/macros/s/AKfycbzVnDWxPtew1hFOfx6IiaCXZoirqNJS802LywmHoAdtsIl9WCjAwDAaVIlteh56dV2m/exec)
```[cite: 2] 
2. Complete all required documentation fields:
   - Project Title & Author Name
   - Raw Dataset Description & Initial Quality Issues Identified
   - Cleaning Steps Executed (Deduplication, TRIM, Case Fixes)
   - Formulas & Pivot Metrics Created (SUMIF, XLOOKUP, Pivot Matrices)
   - Visual Insights & Data Recommendations
3. Click 'Generate README' and download/copy your formatted README.md text file.
```

#### 📅 DAY 12 (Unit 6, Part B): GitHub Deployment & Final Showcase (60–90 Mins)
* **Focus:** GitHub Repository Creation, Asset Upload, & Final Showcase
* **Lab:** Deployment & Class Presentation
* [Readme Code Generator](https://script.google.com/macros/s/AKfycbzVnDWxPtew1hFOfx6IiaCXZoirqNJS802LywmHoAdtsIl9WCjAwDAaVIlteh56dV2m/exec)
```[cite: 2]
```text
[EXERCISE 6.2] GitHub Portfolio Publishing
1. Log into GitHub -> Click 'New Repository'.
2. Name repository: 'excel-data-analytics-portfolio'. Set visibility to Public.
3. Upload project assets:
   - Final Cleaned Excel Workbook (.xlsx)
   - Raw Data Source (.csv)
   - Generated Project Documentation (README.md)
   - Dashboard Screenshot (PNG/JPG)
4. Copy public repository link for final course submission and peer presentations!
```
* **Unit 6 Deliverable:** A live public GitHub portfolio repository containing your workbook, documentation, raw data, and dashboard screenshot.

---

## 📬 Instructor Contact Information

If you have questions regarding the course, daily schedules, or student progress, please reach out:

* ✉️ **Email:** marlainna456@gmail.com  
* 💼 **Office Hours:** Available by appointment after class hours  
```

