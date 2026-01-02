import pandas as pd

# read cancel order excel
cancel_df = pd.read_excel("cancel order.xlsx", header = None)

# skip the headline
cancel_df = cancel_df.iloc[1:].reset_index(drop = True)
col0 = cancel_df[0].tolist()

# each product occupys every 2 or 4 lines
records = []
i = 0
while i < len(col0):
    upc = str(col0[i]).strip() if i < len(col0) else ""
    product_name = str(col0[i + 1]).strip() if i + 1 < len(col0) else ""
    possible_shelf_or_stock = str(col0[i + 2]).strip() if i + 2 < len(col0) else ""

    if possible_shelf_or_stock.isnumeric():
        # what if there is no shelf code
        shelf_code = None
        canceled_qty = int(float(possible_shelf_or_stock))
        i += 3
    else:
        # what if there is shelf code
        shelf_code = possible_shelf_or_stock
        canceled_qty = 0
        if i + 3 < len(col0):
            try:
                canceled_qty = int(float(str(col0[i + 3]).strip()))
            except:
                canceled_qty = 0
        i += 4

    records.append({
        "UPC": upc,
        "Product Name": product_name,
        "Shelf Code": shelf_code,
        "Canceled Qty": canceled_qty
    })

# the result
cancel_result = pd.DataFrame(records)

# save the result to excel
cancel_result.to_excel("parsed_cancel_order.xlsx", index = False)
