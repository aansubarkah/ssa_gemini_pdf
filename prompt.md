# Prompt: Convert Store.pdf to Excel using Python (Gemini Sandbox)

> **Purpose:** Ask Gemini to write and run Python code that extracts the table from `Store.pdf` and saves it as `Store.xlsx`.
> **Audience:** Students / non-programmers. Tweak only the sections marked **(EDITABLE)**.
> **Important:** This version is intentionally free of Python syntax. The user does not need to understand Python; they only need to describe what they want in plain language.

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

## 3. INSTRUCTIONS FOR GEMINI (Plain-language steps)

### 3.1 Setup

Gemini will use its sandboxed Python environment. Ask it to install any missing tools it needs, then write and run the script itself.

### 3.2 Read every page

- Read the entire PDF from page 1 to page 15.
- Identify the table on each page.

### 3.3 Clean the data

- Use the first row of page 1 as the column header.
- Skip the first row on every later page, so the header appears only once.
- Combine all other rows into one continuous list.
- If a piece of information (especially `rowguid` or `Demographics`) is broken into two or more cells on the same row, put it back together into one value.
- Remove fully blank rows.

### 3.4 Create the Excel file

- Create an Excel workbook.
- Put the combined table into the first sheet.
- Save the file as `Store.xlsx`.

### 3.5 Verify and report

After saving, ask Gemini to show:

- How many data rows it extracted (excluding the header).
- The first 5 rows as a preview.
- Confirmation that `Store.xlsx` was created.
- Any warnings about missing values or rows that looked messy.

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
| Excel file columns look too narrow                     | Default column widths                                             | "Auto-fit column widths in the Excel file after saving"               |

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

Please use Python to:
1. Install any missing tools you need.
2. Read every page of Store.pdf.
3. Extract the table from each page.
4. Use the first row of page 1 as the column header.
5. Skip the header row on all later pages.
6. Merge any split cells on each row, especially in rowguid and Demographics, so every row has exactly 6 values matching the header.
7. Save the combined table as Store.xlsx.
8. Tell me how many data rows were extracted, show the first 5 rows, confirm that Store.xlsx was created, and warn me about any messy rows.

Finally, provide the Store.xlsx file so I can download it.
```

---

## 8. STUDENT NOTES

- **Do not change section 7 unless the output from Gemini is wrong.**
- If the output is wrong, first check section 5 (Troubleshooting).
- If a fix is needed, add one extra sentence at the end of the prompt in section 7 explaining the problem, e.g.:
  - `"The rowguid values are split into two cells; merge them."`
  - `"Some rows have 7 columns instead of 6; remove the extra empty cell."`
