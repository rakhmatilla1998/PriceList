import openpyxl
import json
import requests
import os
import pandas as pd

# wb_obj = openpyxl.load_workbook("demo.xlsx")
# sheet_obj = wb_obj.active
#
# #Getting the max number of rows
# max_rows = len(sheet_obj['A'])

price_list_df = pd.read_excel("demo.xlsx")
#print(price_list_df)

#iterating over a dataframe
for i in range(len(price_list_df)):
    item_name = price_list_df.iloc[i,0]
    price_list = [
    {"PriceList": 1, "Price": price_list_df.iloc[i,1]},
    {"PriceList": 2, "Price": price_list_df.iloc[i,2]},
    {"PriceList": 3, "Price": price_list_df.iloc[i,3]},
    {"PriceList": 4, "Price": price_list_df.iloc[i,4]}
    ]
    item_prices = {"ItemPrices": price_list}
    print(item_prices)
