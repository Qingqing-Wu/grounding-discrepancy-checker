# Grounding Discrepancy Checker

A Python-based desktop application for detecting grounding discrepancies across
multiple warehouse operational data sources.

The tool automates a manual, Excel-heavy reconciliation process by applying
rule-based matching logic and generating discrepancy reports and summaries.

---

## 30-Second Overview

Grounding discrepancies occur when physical inventory grounding records do not
match supporting operational data such as inbound receipts, value-added records,
cancel orders, or transfer orders.

This project provides a **Tkinter-based desktop application** that allows users to
load multiple Excel files, automatically detect mismatches at the UPC level, and
generate discrepancy summaries with optional historical tracking.

The application is designed for daily operational use and does **not require
command-line interaction**.

> **Note:** No raw operational data are included in this repository.  
> The project focuses on discrepancy logic, preprocessing rules, and application design.

![Grounding Discrepancy Demo](docs/demo.gif)

---

## Business Context

Warehouse operations typically generate multiple Excel exports from different
systems, including:

- Grounding records
- Inbound receiving records
- Value-added / customization records
- Cancel order records
- Transfer order records

Because these files are produced independently, mismatches in quantity, timing,
or UPC mapping are common. Manual reconciliation across files is time-consuming
and error-prone, especially at scale.

---

## Project Objectives

- Automate grounding discrepancy detection across multiple Excel sources
- Identify missing or mismatched records at the UPC level
- Reduce manual Excel reconciliation effort
- Provide repeatable, rule-based discrepancy checks
- Support optional historical discrepancy tracking

---

## Key Features

- **Multi-source Excel ingestion**  
  Reads grounding, inbound, value-added, cancel, and transfer Excel files.

- **UPC-level discrepancy logic**  
  Detects missing records and quantity mismatches.

- **Optional historical tracking**  
  Supports cumulative discrepancy tracking via `historical.xlsx`.

- **Desktop GUI (Tkinter)**  
  Simple file selection and execution without CLI usage.

- **Rule-based, explainable logic**  
  All discrepancy checks are implemented explicitly in Python code.

---

## Supported Input Files

The application supports the following Excel inputs (all optional, depending on use case):

- `Grounding.xlsx`
- `InboundRecordFile.xlsx`
- `ExportCustomRecordsFile.xlsx`
- `Cancel Order.xlsx`
- `Transfer Order.xlsx`
- `historical.xlsx` (optional, for cumulative tracking)

Expected column formats are defined in the preprocessing scripts.

---

## How to Run

### 1. Install Dependencies

pip install -r requirements.txt

### 2. Launch the Application

python Auto_Grounding_Discrepancy_Checker_app_2_0.py

## Code Structure

```text
grounding-discrepancy-checker/
├── Auto_Grounding_Discrepancy_Checker_app_2_0.py
├── grounding preprocessing.py
├── cancel order preprocessing.py
├── requirements.txt
├── docs/
│   └── demo.gif
└── README.md

