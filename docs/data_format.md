# Data Format Description

## Grounding.xlsx
Each product is recorded in 3 consecutive rows:
1. UPC
2. Product Name
3. Shelf Code, Grounding Quantity, Operator, Operation Time

## InboundRecordFile.xlsx
- Columns include: UPC, Inbound Qty
- One row per inbound transaction

## ExportCustomRecordsFile.xlsx
- Records value-added and outbound data
- Rows without outbound SN are treated as value-added inbound

## Cancel Order.xlsx
Cancel order records use a variable-length structure:
- UPC
- Product Name
- Either:
  - Canceled Quantity (numeric), or
  - Shelf Code + Canceled Quantity

The preprocessing logic detects the structure automatically.

## Transfer Order.xlsx
- UPC
- Transfer Quantity
