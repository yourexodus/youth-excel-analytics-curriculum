# 🧹 Master Data Cleaning Guide & Excel Answer Key

This guide provides step-by-step Excel cleaning instructions for all 5 sample datasets included in the **Excel Data Analytics Course**. Students can use these step-by-step guides during **Unit 2 (Organizing and Cleaning Data)** and **Unit 3 (Data Detective)**.

---

## 🎮 Dataset 1: Gamer Gear & Tech Store Sales (`01_gamer_gear_sales_messy.csv`)

### 🔍 Identified Data Quality Issues
1. **Duplicate Entries:** Duplicate transactions exist for IDs `GG-1002` and `GG-1008`.
2. **Inconsistent Region Casing:** "NORTH" is capitalized while others are Proper Case ("North").
3. **Inconsistent Date Formats:** Mixed formats like `2026-01-15` vs `01/18/2026`.
4. **Text in Currency Columns:** `Price_Each` contains dollar signs (`$49.99`) mixed with raw numbers (`49.99`), causing Excel to treat some cells as text.
5. **Missing Values:** `Quantity` contains missing values (`NaN`).

### 🛠️ Step-by-Step Excel Cleaning Instructions

#### Step 1: Remove Duplicate Transactions
1. Press `CTRL + A` to select all data.
2. Go to the **Data** tab on the Ribbon.
3. Click **Remove Duplicates**.
4. Keep all columns checked and click **OK**.
   * *Result:* 2 duplicate rows removed (Rows `GG-1002` and `GG-1008`).

#### Step 2: Standardize Region Labels
1. Click the dropdown arrow on the **Region** column header (Column D).
2. Uncheck all, then select only **NORTH**.
3. Overwrite the cells by typing **North** and dragging down or copy-pasting across those rows.
4. Clear the filter to restore the full view.

#### Step 3: Standardize Date Formatting
1. Select the entire **Transaction_Date** column (Column B).
2. Right-click the selected column and select **Format Cells** (or press `CTRL + 1`).
3. Under the **Number** tab, select **Date**.
4. Choose `YYYY-MM-DD` (e.g., `2026-03-14`) and click **OK**.

#### Step 4: Clean Currency Values
1. Select the **Price_Each** column (Column F).
2. Press `CTRL + H` to open the **Find and Replace** dialog.
3. In **Find what**, type: `$`
4. Leave **Replace with** completely blank.
5. Click **Replace All**.
6. With the column still selected, click the **Home** tab -> change the Data Type dropdown from *General* to **Currency** (`$`).

#### Step 5: Handle Missing Values & Add Total Revenue Column
1. Filter the **Quantity** column (Column E) to show empty/blank cells.
2. Replace blanks with `1` (or the class default assumption).
3. In Column H (Row 1), type the header: `Total_Revenue`.
4. In cell `H2`, type the formula:
   ```excel
   =E2*F2
   ```
5. Double-click the bottom-right corner of cell `H2` to fill the formula down for all rows.

---

## 🏆 Dataset 2: Esports Tournament & Player Stats (`02_esports_player_stats.csv`)

### 🔍 Identified Data Quality Issues
1. **Leading/Trailing Whitespace:** Gamer tags like `"PixelNinja "` contain extra hidden spaces.
2. **Duplicate Player Rows:** Player `PRO-1003` (`SlayerX`) appears twice.
3. **Formatting & Decimal Precision:** K/D Ratios have inconsistent decimal places (`2.857` vs `1.5`).
4. **Inconsistent Region Codes:** Region listed as `"na"` instead of `"NA"`.

### 🛠️ Step-by-Step Excel Cleaning Instructions

#### Step 1: Remove Hidden Extra Spaces
1. Insert a new temporary column next to **Gamer_Tag** (Column B). Name it `Clean_Gamer_Tag`.
2. In cell `C2`, enter the `TRIM` formula:
   ```excel
   =TRIM(B2)
   ```
3. Copy cell `C2` down the entire column.
4. Select Column C -> Copy (`CTRL + C`) -> Right-click Column B -> **Paste Values** (`V`).
5. Delete the temporary Column C.

#### Step 2: Remove Duplicate Player IDs
1. Select the dataset (`CTRL + A`).
2. Go to **Data** -> **Remove Duplicates**.
3. Select **Player_ID** as the key column to remove duplicate player entries.

#### Step 3: Calculate K/D Ratio & Set Decimal Precision
1. In cell `H2` (K/D Ratio column), calculate the accurate ratio:
   ```excel
   =F2/G2
   ```
2. Copy the formula down for all players.
3. Select Column H -> Go to **Home** tab -> Number group -> Click **Decrease Decimal** until all values display exactly 2 decimal places (e.g., `2.86`).

#### Step 4: Fix Lowercase Region Codes
1. Select Column D (**Region**).
2. Press `CTRL + H` (Find and Replace).
3. Set **Find what:** `na` and **Replace with:** `NA`. Ensure **Match case** is checked so it doesn't alter other text. Click **Replace All**.

---

## 👟 Dataset 3: Sneaker & Streetwear Resale Tracker (`03_sneaker_streetwear_resale.csv`)

### 🔍 Identified Data Quality Issues
1. **Mixed Data Types in Size:** Numeric shoe sizes (`10.5`) mixed with clothing sizes (`M`, `L`, `XL`), making sorting messy.
2. **Currency Formatting:** `Retail_Price` and `Resale_Price` contain unformatted numbers without currency signs.
3. **Missing Profit Calculation:** No dedicated column calculating profit margins.

### 🛠️ Step-by-Step Excel Cleaning Instructions

#### Step 1: Format Currency Fields
1. Select both **Retail_Price** (Column F) and **Resale_Price** (Column G).
2. Go to **Home** -> Change dropdown to **Currency** (`$`).

#### Step 2: Calculate Estimated Profit & Profit Percentage
1. In Column H (Row 1), label the header: `Estimated_Profit`.
2. In cell `H2`, enter:
   ```excel
   =G2-F2
   ```
3. In Column I (Row 1), label the header: `Profit_Margin_%`.
4. In cell `I2`, enter:
   ```excel
   =H2/F2
   ```
5. Format Column I as **Percentage** (`%`) with 1 decimal place (`Home -> Percentage`).

#### Step 3: Separate Apparel from Sneakers
1. Apply filters (`CTRL + SHIFT + L`).
2. Filter **Category** by `Sneakers` to analyze footwear sizes separately from `Streetwear` apparel sizes (`M`, `L`, `XL`).

---

## 🍿 Dataset 4: Streaming Movies & Anime Analytics (`04_streaming_movies_shows.csv`)

### 🔍 Identified Data Quality Issues
1. **Outlier / Erroneous Data:** Movie `STR-1002` lists an IMDb rating of `1.6` instead of `8.6`.
2. **Text Casing Inconsistencies:** Genre labels mixed like `"anime"` vs `"Anime"`.
3. **Unformatted Numbers:** `Total_Hours_Watched` contains large raw numbers without comma separators (e.g., `286000000`).

### 🛠️ Step-by-Step Excel Cleaning Instructions

#### Step 1: Spot & Fix Rating Outliers
1. Select the **IMDb_Rating** column (Column F).
2. Go to **Home** -> **Conditional Formatting** -> **Highlight Cells Rules** -> **Less Than...**
3. Enter `4.0` and select Red Fill.
4. Inspect identified rows: Row `STR-1002` (*Demon Slayer*) was incorrectly keyed as `1.6`. Correct it to `8.6`.

#### Step 2: Standardize Genre Capitalization
1. Filter Column D (**Genre**).
2. Select lowercase `"anime"` and change all entries to Proper Case `"Anime"`.

#### Step 3: Apply Comma Formatting to Large Numbers
1. Select **Total_Hours_Watched** (Column G).
2. Click **Home** -> Number Group -> Click the **Comma Style** button (`,`) to format numbers clearly (e.g., `286,000,000`).

---

## ⚡ Dataset 5: High School Athletic Combine (`05_high_school_fitness_combine.csv`)

### 🔍 Identified Data Quality Issues
1. **ALL-CAPS Names:** Names typed in uppercase (`MARCUS VANCE`) instead of Proper Case.
2. **Missing Data in Optional Tests:** Non-football athletes missing `Bench_Press_Reps` (shows `NaN` or blank).
3. **GPA Outliers:** Invalid GPA entries exceeding `4.0` (e.g., `45.0` instead of `4.5` or `3.5`).

### 🛠️ Step-by-Step Excel Cleaning Instructions

#### Step 1: Fix ALL-CAPS Names using `PROPER()`
1. Insert a blank column next to **Athlete_Name** (Column B). Label it `Clean_Name`.
2. In cell `C2`, enter:
   ```excel
   =PROPER(B2)
   ```
3. Copy down the column.
4. Copy Column C and **Paste Values** back into Column B. Delete temporary Column C.

#### Step 2: Fix GPA Data Errors
1. Filter **GPA** (Column H) to find values greater than `4.0`.
2. Locate `45.0` for athlete `ATH-1002` and correct it to `3.5` (or correct decimal place).

#### Step 3: Handle Missing Bench Press Values
1. Filter **Bench_Press_Reps** (Column G) by **Blanks** / `NaN`.
2. Replace blank cells with `N/A` or `0` depending on combine evaluation standards so calculations aren't disrupted.

---

## 📋 Summary Checklist for Students (README Submission)

When submitting their final `README.md` file for any of these datasets, students should include:

- [ ] **Original File Name:** (e.g., `01_gamer_gear_sales_messy.csv`)
- [ ] **Cleaned File Name:** (e.g., `01_gamer_gear_sales_CLEANED.xlsx`)
- [ ] **Summary of Issues Found:** List of errors identified during inspection.
- [ ] **Tools & Formulas Used:** List functions used (e.g., `TRIM`, `PROPER`, `REMOVE DUPLICATES`, `XLOOKUP`).
- [ ] **Final Record Count:** Total number of rows remaining after cleaning.