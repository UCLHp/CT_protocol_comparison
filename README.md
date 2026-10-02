# Universal XML Change Detector

A Python tool for comparing two folders of XML files and reporting meaningful differences.

The comparison is designed to work with general XML structures rather than a specific XML schema.

## Features

- Recursively compares XML files in **OLD** and **NEW** folders.
- Detects:
  - Added values
  - Removed values
  - Changed values
  - Added or removed XML files
- Converts XML content into contextual **path → value** pairs.
- Handles repeated XML elements by identifying them using their own attributes or child values rather than relying on element position.
- Treats numerically equivalent values such as `28`, `28.0`, and `28.000` as equal.
- Produces an Excel report of detected differences.

## Versions

### `xml_compare.py`

Basic comparison version.

Uses Python's standard XML parser and produces an Excel report containing:

- Differences
- Identity rules used for repeated XML elements

### HTML report version

The extended version uses `lxml` together with `html_report.py`.

In addition to the Excel report, it generates a browser-based side-by-side XML comparison with:

- OLD and NEW XML views
- highlighted changes
- line numbers
- synchronised scrolling
- next/previous change navigation
- option to show or hide unchanged lines

The XML is pretty-printed before display so the report has consistent indentation and matching line numbers.

## Requirements

Typical dependencies include:

```text
pandas
openpyxl
lxml
```

`tkinter` is used for selecting the OLD, NEW, and output folders.

## Usage

Run the relevant comparison script:

```bash
python xml_compare.py
```

Select:

1. OLD XML folder
2. NEW XML folder
3. Output folder

The comparison report will then be generated in the selected output folder.