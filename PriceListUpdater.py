import openpyxl
import json
import simplejson as s_json
import requests
import os
import pandas as pd

try:
    #reading a credentials file to get credentials
    with open(f"{os.path.dirname(os.path.realpath(__file__))}/credentials") as json_file:
        data = json.load(json_file)
        address = data['address']
        port = data['port']
        db = data['CompanyDB']
        user = data['Username']
        pas = data['Password']
        baseAddress = f"{address}:{port}"
        credentials = {"CompanyDB": db, "Password": pas, "UserName": user}

    with requests.Session() as s:
        print("Подключение к SAP...")
        response = s.post(f"{baseAddress}/b1s/v2/Login", json = credentials, verify = False)
        if(response.status_code == 200):

            #loading data from excel file into a dataframe
            price_list_df = pd.read_excel("Narxlar.xlsx", sheet_name=1, header = None)

            #iterating over a dataframe
            for i in range(len(price_list_df)):
                item_name = str(price_list_df.iloc[i,2])
                price_list = [
                {'PriceList': 1, 'Price': float(price_list_df.iloc[i,5])},
                {"PriceList": 2, "Price": float(price_list_df.iloc[i,5])},
                {"PriceList": 3, "Price": float(price_list_df.iloc[i,6])},
                {"PriceList": 4, "Price": float(price_list_df.iloc[i,7])}
                ]
                item_prices = s_json.dumps({'ItemPrices': price_list}, ignore_nan=True)
                response = s.patch(f"{baseAddress}/b1s/v2/Items('{item_name}')", data = item_prices, verify = False)


except FileNotFoundError:
    print("Не удалось найти файл credentials.")
except requests.exceptions.ConnectionError:
    print("Не удалось подключиться к серверу. Сервер не отвечает.")
except Exception as e:
    print(str(e))
    print("Неизвестная ошибка.")
finally:
    os.system("exit")
