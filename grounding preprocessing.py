import pandas as pd

# read Excel
file_path = "Grounding.xlsx"
df = pd.read_excel(file_path)

products = []

for i in range(0, len(df), 3):
    group = df.iloc[i:i + 3]
    if len(group) == 3:
        upc = str(group.iloc[0, 0])  # line 1: UPC
        product_name = str(group.iloc[1, 0])  # line 2: product name
        shelf_code = group.iloc[2, 0]  # line 3 col 0：Shelf Code
        grounding_num = group.iloc[2, 1]  # line 3 col 1：Grounding Num.
        operator = group.iloc[2, 2]  # line 3 col 2：Operator
        operation_time = group.iloc[2, 3]  # line 3 col 3：Operation Time

        try:
            grounding_num = int(float(grounding_num))
        except:
            grounding_num = None

        product = {
            'UPC': upc,
            'Product Name': product_name,
            'Shelf Code': shelf_code,
            'Grounding Num.': grounding_num,
            'Operator': operator,
            'Operation Time': operation_time
        }
        products.append(product)

# turn into DataFrame
cleaned_df = pd.DataFrame(products)

# save to Excel file
cleaned_df.to_excel("Cleaned_Grounding_Data.xlsx", index=False)


