# CT Protocol Comparison

A **Python tool** for working with **CT protocol HTML exports**.

It supports two workflows:

- **Extract mode**: convert **one HTML file** into a structured **Excel summary**
- **Compare mode**: compare **two HTML files** (**BEFORE** and **AFTER**) and generate:
  - a **comparison report**
  - a highlighted **BEFORE** spreadsheet
  - a highlighted **AFTER** spreadsheet

The aim is to make CT protocol data **easier to read, filter, review, and compare** than in the original HTML format.

---

## Features

- Parses **protocol**, **acquisition**, and **result** sections from CT protocol HTML files
- Converts extracted data into a structured **Excel spreadsheet**
- Supports **single-file extraction**
- Supports **two-file comparison**
- Detects:
  - **removed rows**
  - **added rows**
  - **changed parameter values**
  - **new parameter headers** that appear only in one file
- Produces **colour-highlighted Excel outputs** for visual review
- Automatically **adjusts Excel column widths**
- Uses a simple **Tkinter pop-up interface** for file selection

---

## Requirements

Please install the packages listed in `requirements.txt`.

The script currently uses:

- `pandas`
- `beautifulsoup4`
- `openpyxl`
- `tkinter` *(usually included with standard Python installations)*
- `os` *(standard library)*

---

## How It Works

The script reads CT protocol HTML content and extracts data from:

- **protocol headings**
- **acquisition labels**
- **result labels**
- **parameter / value tables**

Each extracted row is stored with the following **identifier columns**:

- `Protocol`
- `Acquisition Number`
- `Label`
- `Type`
- `Result Label`

These columns are used as a **composite key** to identify matching rows when comparing two files.

In **Compare mode**, the script classifies differences as:

- **Removed**: present in **BEFORE** but not in **AFTER**
- **Added**: present in **AFTER** but not in **BEFORE**
- **Changed**: same row exists in both, but one or more parameter values differ

The comparison checks the **union of parameter columns from both files**, so it can also detect **new parameter headers** introduced in the newer file.

---

## How to Use

### 1. Run the script

Run:

```bash
python ct_protocol_summary.py