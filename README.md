# Grounding Discrepancy Checker 2.0

A Python desktop tool (Tkinter) to detect discrepancies between Grounding records and aggregated inbound-related sources (Inbound / Value-Add / Cancel / Transfer), with optional cumulative history and daily mismatch statistics + charts.

## Features
- Multi-source inbound aggregation
- UPC-level discrepancy detection (Missing Inbound / Missing Grounding / Quantity Mismatch)
- Optional cumulative historical tracking (historical.xlsx)
- Daily discrepancy statistics and plotting

## Input Files (Excel)
Place these files next to `app.py` OR use the Browse buttons:
- Grounding.xlsx (optional)
- InboundRecordFile.xlsx (optional)
- ExportCustomRecordsFile.xlsx (optional)
- Cancel Order.xlsx (optional)
- Transfer Order.xlsx (optional)
- historical.xlsx (optional)

## Run
```bash
pip install -r requirements.txt
python app.py
