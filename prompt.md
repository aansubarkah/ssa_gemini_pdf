# Prompt: Convert Store.pdf to Excel using Python (Gemini Sandbox)

> **Purpose:** Ask Gemini to write and run Python code that extracts the table from `Store.pdf` and saves it as `Store.xlsx`.
> **Audience:** Students / non-programmers. Tweak only the sections marked **(EDITABLE)**.

---

## 1. GOAL (What we want)

Extract the tabular store data from the attached `Store.pdf` and produce a clean Excel file named `Store.xlsx`.

- **Input file:** `Store.pdf`
- **Output file:** `Store.xlsx`
- **Expected content:** one worksheet containing all rows from the PDF table(s).
- **Important:**
  - The table is printed across **15 pages**.
  - Each page repeats the **same 6 columns**:
    1. `BusinessEntityID`
    2. `Name`
    3. `SalesPersonID`
    4. `Demographics`
    5. `rowguid`
    6. `ModifiedDate`

---

## 2. DATA OVERVIEW (What the table looks like)

```
BusinessEntityID | Name                    | SalesPersonID | Demographics | rowguid | ModifiedDate
292              | Next-Door Bike Store    | 279           | ...          | ...     | ...
294              | Professional Sales and  | 276           | ...          | ...     | ...
```

- The first row of each page is the **header**. Do **not** duplicate it in the final Excel file.
- The `Demographics` column contains long XML text.
- The `rowguid` is a UUID but may be split across table cells; please merge split cells back into one value per logical row.

---

## 3. INSTRUCTIONS FOR GEMINI

Generate Python code that does the following, then run it in the sandbox and provide the resulting `Store.xlsx` file for download.

### 3.1 Load libraries

Use:

```python
import pdfplumber
import pandas as pd
```

If these packages are not installed, install them with:

```python
!pip install pdfplumber pandas openpyxl
```

### 3.2 Read every page

1. Open `Store.pdf` with `pdfplumber`.
2. Loop through **all pages**.
3. On each page, extract the table using `page.extract_tables()`.

### 3.3 Clean the data

1. Skip empty tables.
2. Treat the **first non-empty row** of the first page as the header.
3. For every other page, skip its header row.
4. Combine all remaining rows into a single list.
5. If any logical cell (especially `rowguid` or `Demographics`) is split across multiple cells on the same row, concatenate the pieces with no spaces unless a space is already present.
6. Remove fully blank rows.

### 3.4 Create the Excel file

1. Convert the combined data into a `pandas.DataFrame` using the header from step 3.3.2.
2. Save it as `Store.xlsx` with `index=False`:

```python
df.to_excel("Store.xlsx", index=False, engine='openpyxl')
```

### 3.5 Verify and report

After saving, print:

- The number of rows extracted (excluding the header).
- The first 5 rows of the DataFrame.
- A confirmation that `Store.xlsx` was created.
- Any warnings about missing values or split cells that could not be merged cleanly.

---

## 4. EXPECTED OUTPUT FILES

| File          | Description                                          |
|---------------|------------------------------------------------------|
| `Store.xlsx`  | Clean Excel workbook with one sheet: all store rows  |

---

## 5. TROUBLESHOOTING GUIDE (Edit if new problems appear)

| Symptom                                                | Likely cause                                                      | What to ask Gemini to fix                                            |
|--------------------------------------------------------|-------------------------------------------------------------------|-----------------------------------------------------------------------|
| Header appears in the middle of the data               | Page header not being skipped                                     | "Skip the first row on every page after page 1"                       |
| `rowguid` looks broken into several columns            | PDF wraps one logical cell across cells                           | "Concatenate the rowguid pieces for each row into one column"         |
| Some rows have the wrong number of columns             | Page contains extra text lines that were mistaken for table rows  | "Only keep rows that have 6 non-empty values"                         |
| Duplicate rows appear                                  | Same table printed on every page and headers are kept twice       | "Remove duplicate rows, and keep only one header at the top"          |
| Excel file columns look too narrow                     | Default column widths                                             | "Auto-fit column widths in openpyxl after saving"                     |

---

## 6. OPTIONAL ADVANCEMENTS (Only if needed)

- **Parse `Demographics` XML** into separate columns such as `AnnualRevenue`, `BankName`, `BusinessType`, `YearOpened`, `Specialty`, `SquareFeet`, `Brands`, `Internet`, `NumberEmployees`.
- **Format** `ModifiedDate` as a proper date column.
- **Rename the worksheet** from `Sheet1` to `Stores`.

---

## 7. PROMPT TO PASTE INTO GEMINI (Copy from here to the chat)

```text
You have a sandboxed Python environment and a file called Store.pdf attached to this chat.

The PDF is 15 pages long. Each page contains the same store table with these 6 columns:
BusinessEntityID, Name, SalesPersonID, Demographics, rowguid, ModifiedDate.

Please write and execute Python code that:
1. Installs pdfplumber, pandas, and openpyxl if they are not already available.
2. Reads every page of Store.pdf.
3. Extracts tables from each page using pdfplumber.
4. Uses the first page's header row as the column names.
5. Skips the header row on all later pages.
6. Merges any split cells on each row, especially rowguid and Demographics, so every row has exactly 6 logical values matching the header.
7. Creates a DataFrame and saves it as Store.xlsx with openpyxl (index=False).
8. Prints the number of data rows, the first 5 rows, a confirmation that Store.xlsx was created, and any warnings about messy rows.

Finally, provide the Store.xlsx file so I can download it.
```

---

## 8. STUDENT NOTES

- **Do not change section 7 unless the output from Gemini is wrong.**
- If the output is wrong, first check section 5 (Troubleshooting).
- If a fix is needed, add one extra sentence at the end of the prompt in section 7 explaining the problem, e.g.:
  - `"The rowguid values are split into two cells; merge them."`
  - `"Some rows have 7 columns instead of 6; remove the extra empty cell."`
