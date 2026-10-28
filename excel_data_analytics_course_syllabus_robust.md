# 📊 Excel Data Analytics Course Syllabus

> **Target Audience:** High School Students (Ages 15–18)  
> **Pacing:** 6 Days (60–90 Minutes per Day / Unit)  
> **Prerequisites:** Windows 10/11 (64-bit), Microsoft Excel, Google Chrome  

---

## 🎨 Course Overview & Core Values

Welcome to **Data Analytics**! In this hands-on course, you'll discover how to turn raw, messy datasets into actionable intelligence using Microsoft Excel. You'll master data hygiene, dynamic formulas, conditional logic, PivotTable summaries, and interactive dashboards, culminating in a professional data portfolio.

| 🌟 **Our Core Values** | 🎯 **What You Will Learn** | 💻 **System Requirements** |
| :--- | :--- | :--- |
| 🟢 **Respectful** & Supportive | 🔍 Connect & organize data sources | **OS:** Windows 10, 11, or Server 2016+ (64-bit) |
| 💡 **Curious** & Engaged | 🧹 Identify & correct data quality issues | **CPU:** 1 GHz or faster 64-bit (x64) |
| 🛡️ **Safe for Learning** | ⚡ Transform & combine data using Power Query | **RAM:** 2 GB min (4 GB recommended) |
| 🚀 Tech Progress & Growth | 📈 Build a data preparation portfolio | **Display:** 1440x900 or 1600x900 (100% scale) |

---

## 📁 Course Datasets & Interactive Tools

Students select 1 primary dataset from the local `data/` folder to analyze across all 6 units:

* 🎮 **Dataset 1:** [Gamer Gear & Tech Store Sales (`data/01_gamer_gear_sales_messy.csv`)](data/01_gamer_gear_sales_messy.csv)
* 🏆 **Dataset 2:** [Esports Tournament & Player Stats (`data/02_esports_player_stats.csv`)](data/02_esports_player_stats.csv)
* 👟 **Dataset 3:** [Sneaker & Streetwear Resale Tracker (`data/03_sneaker_streetwear_resale.csv`)](data/03_sneaker_streetwear_resale.csv)
* 🍿 **Dataset 4:** [Streaming Movies & Anime Analytics (`data/04_streaming_movies_shows.csv`)](data/04_streaming_movies_shows.csv)
* ⚡ **Dataset 5:** [High School Athletic Combine Performance (`data/05_high_school_fitness_combine.csv`)](data/05_high_school_fitness_combine.csv)

🛠️ **Course Tools & Web Applications:**
* 🧹 **Data Cleaning Answer Key:** [Master Data Cleaning Guide](data_cleaning_guide.md)
* 📝 **README Generator Web App:** [Build My README App](https://script.google.com/macros/s/AKfycbzVnDWxPtew1hFOfx6IiaCXZoirqNJS802LywmHoAdtsIl9WCjAwDAaVIlteh56dV2m/exec)

---

## 📅 Robust Day-by-Day Unit Schedule & In-Depth Exercises

---

### 🟢 DAY 1: Unit 1 — Excel Navigation, Data Architecture & Data Type Audits
* **Duration:** 60–90 Minutes  
* **Focus:** Data Inspection, Interface Speed Keys, Data Types, & Grid Structuring  
* **Assessment:** Data Type Audit Log & Navigation Speed Test  

#### 🎯 Learning Objectives
- Understand how modern businesses use data analytics to make operational decisions.
- Navigate large data grids seamlessly using keyboard shortcuts (`CTRL + Arrow Keys`, `CTRL + SHIFT + Down`).
- Differentiate between categorical (text), quantitative (numeric/currency), and temporal (date/time) data types.
- Format columns explicitly to prevent automated date/number conversion errors.

#### ⏱️ Minute-by-Minute Session Structure
* **00–15 Min:** Instructor demo on raw CSV structural errors and data types.
* **15–50 Min:** Guided Exploration Lab 1.1 & 1.2.
* **50–80 Min:** Independent Data Detective Challenge 1.3.
* **80–90 Min:** Student reflection log & exit ticket.

#### 🧪 Day 1 Hands-On Labs & Exercises

```text
[EXERCISE 1.1] Grid Navigation & Boundary Inspection (20 Mins)
1. Launch Excel and open your selected dataset from the 'data/' folder.
2. Click cell A1 and press CTRL + DOWN ARROW to find the exact last row. Record the row count.
3. Press CTRL + RIGHT ARROW to identify the total number of attribute columns.
4. Select Row 1 (Headers) -> Navigate to View Tab -> Freeze Panes -> Click 'Freeze Top Row'.
5. Scroll 100 rows down and confirm headers stay locked on screen.

[EXERCISE 1.2] Data Type Classification & Formatting Audit (20 Mins)
1. Create a 2-column table on Sheet2 labeled 'Column_Name' and 'Identified_Data_Type'.
2. Inspect each column in your dataset and categorize as: [Text, Currency, Decimal, Integer, Date].
3. Select any date columns -> Right-click -> Format Cells -> Select Short Date (YYYY-MM-DD).
4. Select any price/sales columns -> Format as Currency ($) with 2 decimal places.

[EXERCISE 1.3] Data Detective Challenge: Spotting Mixed Data Types (30 Mins)
1. Inspect columns containing numerical values (e.g., Price, GPA, or K/D Ratio).
2. Look for cells where numbers are stored as text (e.g., "$49.99" mixed with "49.99").
3. Use conditional formatting: Home -> Conditional Formatting -> Highlight Cells Rules -> Text That Contains "$".
4. Highlight all dirty cells and list their exact cell coordinates in your audit log.
```

* **Day 1 Exit Deliverable:** A worksheet containing locked headers, formatted numbers/dates, and an audit table listing 3 data-type inconsistencies found.

---

### 🔵 DAY 2: Unit 2 — Data Hygiene, Cleaning Workflows & String Manipulation
* **Focus:** Deduplication, Filter Diagnostics, Whitespace Scrubbing, & Case Standardization  
* **Assessment:** Before/After Data Cleaning Checklist & Log  
* **Duration:** 60–90 Minutes  

#### 🎯 Learning Objectives
- Import raw CSV files into Excel using Power Query / Data Import tools without corrupting formats.
- Identify and remove duplicate records based on primary unique identifiers (IDs).
- Scrub hidden leading and trailing spaces using the `TRIM()` function.
- Standardize inconsistent text casing (`UPPER`, `LOWER`, `PROPER`).

#### ⏱️ Minute-by-Minute Session Structure
* **00–15 Min:** Deep dive into common real-world data errors (duplicates, trailing spaces, case sensitivity).
* **15–55 Min:** Guided Cleaning Workflow (Labs 2.1 & 2.2).
* **55–80 Min:** Data Detective Cleaning Challenge 2.3.
* **80–90 Min:** Peer verification & checklist sign-off.

#### 🧪 Day 2 Hands-On Labs & Exercises

```text
[EXERCISE 2.1] CSV Import & Deduplication (20 Mins)
1. Open a blank Excel workbook -> Go to Data Tab -> Get Data -> From File -> From Text/CSV.
2. Browse to 'data/' and select your dataset file. Set Data Type Detection to 'Based on entire dataset'.
3. Press CTRL + A to highlight the dataset -> Go to Data Tab -> Click 'Remove Duplicates'.
4. Ensure 'My data has headers' is checked. Keep all columns selected and click OK.
5. Record the exact number of duplicate rows removed in your daily log.

[EXERCISE 2.2] String Scrubbing with TRIM() and PROPER() (25 Mins)
1. Insert a blank column next to your text column (e.g., Gamer_Tag, Athlete_Name, Product_Name).
2. Title the new column 'Clean_Text_Value'.
3. In cell C2, enter the formula: =PROPER(TRIM(B2))
4. Double-click the fill handle (bottom right of cell C2) to copy down all rows.
5. Highlight Column C -> Copy (CTRL + C) -> Right-click Column B -> Paste Special -> Values (V).
6. Delete temporary Column C.

[EXERCISE 2.3] Data Detective Challenge: Filter Diagnostics (25 Mins)
1. Highlight headers -> Press CTRL + SHIFT + L to apply AutoFilters.
2. Click the filter arrow on region/category columns (e.g., Region, Genre, Brand).
3. Search for typos, capitalization splits (e.g., "NORTH" vs "North"), or missing values (Blanks/NaN).
4. Use Find and Replace (CTRL + H) to standardize all typos to Proper Case.
5. Filter by 'Blanks' in numeric columns and fill missing values with '0' or 'N/A'.
```

* **Day 2 Exit Deliverable:** A fully cleaned raw data tab with 0 duplicates, trimmed text strings, and standardized categorical filters.

---

### 🟡 DAY 3: Unit 3 — Analytics Formulas, Conditional Logic & Data Lookups
* **Focus:** Summary Aggregations, Math Logic (`IF`, `SUMIF`, `COUNTIF`), and `XLOOKUP`  
* **Assessment:** "Data Detective" Quantitative Analysis Sheet  
* **Duration:** 60–90 Minutes  

#### 🎯 Learning Objectives
- Calculate fundamental statistical metrics using `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, and `COUNTA`.
- Execute conditional aggregations based on single and multiple criteria using `COUNTIF` and `SUMIF`.
- Write logical `IF` statements to flag outliers, high-performers, or missing values.
- Merge disparate tables seamlessly using `XLOOKUP`.

#### ⏱️ Minute-by-Minute Session Structure
* **00–15 Min:** Formula syntax tutorial (relative vs absolute cell referencing `$A$1`).
* **15–50 Min:** Formula Building Lab 3.1 & 3.2.
* **50–80 Min:** Data Detective Logic Challenge 3.3.
* **80–90 Min:** Formula verification audit.

#### 🧪 Day 3 Hands-On Labs & Exercises

```text
[EXERCISE 3.1] Core Summary KPI Metrics (20 Mins)
Create a separate sheet named 'KPI_Summary' and construct the following calculations:
1. Total Revenue / Score Sum:     =SUM(CleanData!E2:E200)
2. Average Metric Value:         =AVERAGE(CleanData!E2:E200)
3. Total Valid Records:          =COUNT(CleanData!E2:E200)
4. Peak Recorded Value:          =MAX(CleanData!E2:E200)
5. Minimum Recorded Value:       =MIN(CleanData!E2:E200)

[EXERCISE 3.2] Conditional Logic & Business Rules (25 Mins)
1. Calculate total sales/scores for a specific category:
   =SUMIF(CleanData!C2:C200, "Apparel", CleanData!E2:E200)
2. Count how many records exceed the average threshold:
   =COUNTIF(CleanData!E2:E200, ">" & B2)
3. Add a calculated column titled 'Performance_Tier' using a nested IF formula:
   =IF(E2>=100, "Top Tier", IF(E2>=50, "Mid Tier", "Low Tier"))

[EXERCISE 3.3] Data Detective Challenge: Multi-Table Lookups (25 Mins)
1. Create a reference table listing Product/Item IDs and their Target Goals on Sheet3.
2. In your main dataset, add a column titled 'Target_Status'.
3. Write an XLOOKUP formula to pull target benchmarks into your main dataset:
   =XLOOKUP(A2, TargetSheet!A$2:A$20, TargetSheet!B$2:B$20, "No Target Found")
4. Identify all records that failed to reach their target goal.
```

* **Day 3 Exit Deliverable:** A completed summary sheet featuring calculated statistical KPIs, `SUMIF`/`COUNTIF` tables, and working `XLOOKUP` functions.

---

### 🟣 DAY 4: Unit 4 — Summarizing Data with PivotTables & Cross-Tabulation
* **Focus:** Dynamic Tables (`CTRL + T`), PivotTable Layouts, Grouping, & Interactive Slicers  
* **Assessment:** PivotTable Diagnostic Report  
* **Duration:** 60–90 Minutes  

#### 🎯 Learning Objectives
- Convert static cell ranges into dynamic structured Excel Tables (`CTRL + T`).
- Aggregate complex datasets instantly using PivotTables without writing complex formulas.
- Group raw dates into months, quarters, and years, or numbers into custom bins.
- Connect dynamic Slicers and Timelines for interactive filtering.

#### ⏱️ Minute-by-Minute Session Structure
* **00–15 Min:** PivotTable mechanics, field lists (Filters, Columns, Rows, Values), and calculation settings.
* **15–50 Min:** PivotTable Creation Lab 4.1 & 4.2.
* **50–80 Min:** Data Detective Pivot Exploration 4.3.
* **80–90 Min:** Student presentations of key insights.

#### 🧪 Day 4 Hands-On Labs & Exercises

```text
[EXERCISE 4.1] Converting Ranges & Generating PivotTables (20 Mins)
1. Click inside your clean data sheet -> Press CTRL + T -> Ensure 'Headers' is checked -> Click OK.
2. Name your table 'CleanDataTable' in the Table Design tab.
3. Click Insert Tab -> PivotTable -> Choose 'New Worksheet'. Name sheet 'Pivot_Analysis'.
4. Drag your primary categorical field (e.g., Brand, Genre, Region, Sport) into ROWS.
5. Drag your main numeric field (e.g., Revenue, Price, Score) into VALUES.
6. Right-click any value cell -> Summarize Values By -> Choose 'SUM' or 'AVERAGE'.

[EXERCISE 4.2] Multi-Dimensional Matrix & Date Grouping (25 Mins)
1. Drag a secondary categorical field into COLUMNS to build a cross-tabulation matrix.
2. If your data includes dates, drag Date to ROWS. Right-click a date cell -> Group -> Select 'Months' & 'Years'.
3. Change Value Field Settings -> Show Values As -> '% of Column Total' to analyze market share percentage.

[EXERCISE 4.3] Data Detective Challenge: Interactive Slicer Dashboard (25 Mins)
1. Click inside your PivotTable -> Go to PivotTable Analyze -> Insert Slicer.
2. Check 2 fields (e.g., Customer_Type, Grade_Level, Rank_Tier, Age_Rating).
3. Insert a Timeline slicer if date fields are present.
4. Format slicers into multi-column layouts and test dynamic interactive filtering.
5. Identify the top 3 revenue/performance categories under specific slicer filter conditions.
```

* **Day 4 Exit Deliverable:** A dedicated `Pivot_Analysis` sheet with at least 2 structured PivotTables, formatted percentages, and interactive Slicers.

---

### 🟧 DAY 5: Unit 5 — Data Visualization, Chart Selection & Dashboard Mockups
* **Focus:** Visual Design Principles, Chart Selection, Custom Formatting, & Storytelling  
* **Assessment:** Chart Rationale Challenge & Static Dashboard  
* **Duration:** 60–90 Minutes  

#### 🎯 Learning Objectives
- Select the correct chart type based on analytical goals: Bar (comparison), Line (time trends), Scatter Plot (correlations), or Column (distribution).
- Format visual components (chart titles, axis labels, gridlines, data labels) for executive presentation.
- Avoid misleading visual practices (truncated Y-axes, 3D charts, chart clutter).
- Assemble multiple visual components into a organized executive dashboard tab.

#### ⏱️ Minute-by-Minute Session Structure
* **00–15 Min:** Data visualization best practices, color theory, and chart selection matrix.
* **15–50 Min:** Chart Construction Lab 5.1 & 5.2.
* **50–80 Min:** Data Detective Dashboard Design 5.3.
* **80–90 Min:** Visual design review.

#### 🧪 Day 5 Hands-On Labs & Exercises

```text
[EXERCISE 5.1] Time-Series Trend & Category Comparison Charts (25 Mins)
1. Select your monthly summary PivotTable -> Click Insert Tab -> Line Chart -> Line with Markers.
2. Change title to descriptive format: "2026 Monthly Trend Analysis".
3. Click Add Chart Element -> Axis Titles -> Add Primary Vertical ("Revenue in USD").
4. Select category total data -> Click Insert Tab -> Clustered Column Chart.
5. Format the top-performing data bar with a bold accent color; color all other bars dark gray.

[EXERCISE 5.2] Scatter Plot Correlation Analysis (20 Mins)
1. Select two numeric columns (e.g., Retail vs Resale Price, 40yd Dash vs Vertical Jump, IMDb Rating vs Hours Watched).
2. Click Insert Tab -> Scatter Chart -> Scatter with Only Markers.
3. Add Axis Titles for both X (Independent variable) and Y (Dependent variable).
4. Right-click data points -> Add Trendline -> Display R-squared value on chart to test correlation strength.

[EXERCISE 5.3] Data Detective Challenge: Executive Dashboard Assembly (25 Mins)
1. Create a blank worksheet named 'Executive_Dashboard'. Turn off Gridlines (View Tab -> Uncheck Gridlines).
2. Cut and paste your Line Chart, Column Chart, Scatter Plot, and Slicers onto this sheet.
3. Align visual components neatly using Align to Grid (Page Layout -> Align).
4. Connect Slicers to all charts: Right-click Slicer -> Report Connections -> Check all PivotTables.
```

* **Day 5 Exit Deliverable:** A clean, ungridlined `Executive_Dashboard` tab with 3 dynamic charts connected to unified Slicers.

---

### 🚀 DAY 6: Unit 6 — Portfolio Assembly & Mandatory README Generator Deployment
* **Focus:** End-to-End Audit, README Generator Web App, GitHub Deployment, & Showcase  
* **Assessment:** Final Project Presentation & Complete GitHub Repository  
* **Duration:** 60–90 Minutes  

#### 🎯 Learning Objectives
- Perform a final end-to-end data audit to verify workbook accuracy and structure.
- Utilize the mandatory custom README Generator web app to create professional Markdown documentation.
- Package project files, raw datasets, cleaned workbooks, and README documentation for submission.
- Publish and present your data analytics portfolio repository.

#### ⏱️ Minute-by-Minute Session Structure
* **00–10 Min:** Overview of portfolio standards and final project submission requirements.
* **10–45 Min:** Mandatory README Documentation Generation (Lab 6.1).
* **45–75 Min:** GitHub Repository Setup & Asset Upload (Lab 6.2).
* **75–90 Min:** Student Project Showcase & Final Evaluation.

#### 🧪 Day 6 Hands-On Labs & Exercises

```text
[EXERCISE 6.1] Mandatory README Documentation Generation (35 Mins)
1. Open the mandatory README Generator web application in your browser:
   👉 https://script.google.com/macros/s/AKfycbzVnDWxPtew1hFOfx6IiaCXZoirqNJS802LywmHoAdtsIl9WCjAwDAaVIlteh56dV2m/exec
2. Complete all interactive form fields detailing your analytical project:
   - Project Title & Author Name
   - Dataset Selected & Raw Source File Path
   - Summary of Data Quality Errors Found (duplicates, bad types, missing data)
   - Step-by-Step Data Cleaning Workflow Executed
   - Excel Formulas & Functions Utilized (e.g., TRIM, PROPER, SUMIF, XLOOKUP)
   - PivotTable Summaries & Chart Selection Rationale
   - Key Business Insights & Data-Driven Recommendations
3. Click 'Generate README' -> Copy or Download the generated README.md text file to your project directory.

[EXERCISE 6.2] GitHub Portfolio Deployment & Submission (30 Mins)
1. Log into your GitHub account -> Click 'New Repository'.
2. Name repository: 'excel-data-analytics-portfolio'.
3. Set visibility to Public -> Click Create Repository.
4. Upload project assets:
   - Final Cleaned Workbook (.xlsx)
   - Raw Dataset CSV File (.csv)
   - Completed Markdown Documentation (README.md)
   - Dashboard Screenshot (PNG/JPG)
5. Copy your public GitHub repository link and submit it for final evaluation!
```

* **Day 6 Exit Deliverable:** A live public GitHub portfolio repository containing your cleaned workbook, raw dataset, dashboard screenshot, and generated `README.md` file.

---

## 📬 Instructor Contact Information

If you have questions regarding the course, daily schedules, or student progress, please reach out:

* ✉️ **Email:** marlainna456@gmail.com  
* 💼 **Office Hours:** Available by appointment after class hours  
```

---

### Key Enhancements Made:
1. **Pacing Realism:** Every day is structured with time blocks (e.g., *00–15 Min demo*, *15–50 Min lab*, *50–80 Min challenge*) ensuring students remain engaged for the full 60–90 minutes.
2. **Technical Depth:** Added specific formula syntax, nested logic (`IF` within `IF`), `XLOOKUP` defaults, string functions (`TRIM`, `PROPER`), and table conversions (`CTRL + T`).
3. **Data Detective Challenges:** Applied practical diagnostic tasks (such as scatter plot trendline analysis, multi-table lookups, and filter audits) so advanced students have deep tasks to explore.
4. **Integration of README Generator Tool:** Seamlessly built your Google Apps Script README generator app directly into Day 6 as the mandatory documentation deliverable.