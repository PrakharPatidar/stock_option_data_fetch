import requests
import datetime
import pandas as pd
from openpyxl import load_workbook
import traceback
import sys

url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
headers = {'User-Agent': 'Mozilla/5.0'}

retry_url_count = 10

def call_api(url, headers):
    return requests.get(url, headers = headers).json()

try:
    api_call_count = 0
    while True:
        try:
            api_data = call_api(url, headers = headers)
        except Exception as e:
            print('Retrying url')
            if api_call_count < retry_url_count:
                api_call_count += 1
                continue
            else:
                print(f'Retried for {retry_url_count}. Exiting now...')
                sys.exit(1)

        break

    imp_data = {
        "OI_Calls": [
        ],
        "Change_In_OI_Calls": [
        ],
        "Strick_Price": [
        ],
        "Change_In_OI_Puts": [
        ],
        "OI_Puts": [
        ]
    }

    for each_data in api_data['filtered']['data']:
        imp_data["OI_Calls"].append(each_data["CE"]["openInterest"])
        imp_data["Change_In_OI_Calls"].append(each_data["CE"]["changeinOpenInterest"])
        imp_data["Strick_Price"].append(each_data["strikePrice"])
        imp_data["Change_In_OI_Puts"].append(each_data["PE"]["changeinOpenInterest"])
        imp_data["OI_Puts"].append(each_data["PE"]["openInterest"])

    excel_file_path = 'output.xlsx'
    sheet_name = datetime.datetime.now().strftime("%d-%m-%Y %H-%M-%S")
    print('sheet name', sheet_name)

    book = load_workbook(excel_file_path)

    writer = pd.ExcelWriter(excel_file_path, engine = 'openpyxl')
    writer.book = book

    to_save_excel = pd.DataFrame.from_dict(imp_data)
    to_save_excel.to_excel(writer, sheet_name)
    writer.save()
    writer.close()
except Exception as e:
    print('Something went wrong')
    print(e)
    traceback.print_exc()