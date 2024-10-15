import pandas as pd
import json
import requests as r
from datetime import datetime,timedelta
import time
import tkinter as tk
from tkinter import filedialog, messagebox

after_json = []
file_path = 'before.xlsx'
column_name_mapping = {
    'שם מלא': 'name',
    'כרטיס אשראי': 'number',
    'תוקף': 'date',
    'cvv' : 'cvv'
}
paymentUrl = "https://live.payme.io/api/pay-sale"
saleUrl = "https://live.payme.io/api/generate-sale"
headers = {
    "Content-Type": "application/json",
    "Accept": "application/json"
}
dummyDataPayment = {
    "currency": "ILS",
    "language": "he",
    "buyer_name": "test name",
    "sale_price": "1",
    "installments": "1",
    "capture_buyer": 0,
    "payme_sale_id": "Copy here payme_sale_id from generate-sale request",
    "credit_card_cvv": 652,
    "credit_card_exp": "0521",
    "sale_return_url": "https://shmoo",
    "seller_payme_id": "MPL17283-73752XFU-BOLZOJAN-3ZDHJCGY",
    "credit_card_number": 4111111111111111,
    "buyer_email": "imcry3@gmail.com"
    }
dummyDataSale = {
    "seller_payme_id":"MPL17283-73752XFU-BOLZOJAN-3ZDHJCGY",
    "sale_payment_method":"credit-card",
    "product_name":"בדיקה",
    "currency":"ILS",
    "sale_price":"1",
    "installments":"1",
    "sale_type":"token",
    "sale_name":"חנה",
    "buyer_perform_validation": True
    }

try:
    df = pd.read_excel(file_path)
except Exception as e:
    print("Error loading Excel file:", e)
    exit()
# Rename columns from Hebrew to English
df.rename(columns=column_name_mapping, inplace=True)

# Convert the DataFrame to a JSON object
data_json = df.to_json(orient='records')
counter = 0
for c in json.loads(data_json):
    customer = {}
    for key,value in c.items():
        if key in column_name_mapping.values():
            customer[key] = value
    try:
        new_date = datetime.fromtimestamp(customer['date']/1000)
        customer['date'] = new_date.strftime("%m%y")
    except:
        print("There is a Type Error on the date")
        c['תושבת הסליקה'] = "There is a Type Error on the date"

    dummyDataSale['sale_name'] = customer['name']
    try:
        sale = r.post(saleUrl,data=json.dumps(dummyDataSale),headers=headers)
    except r.exceptions.HTTPError as e:
        print("Error on api request to generate sale:", e)

    customer['cvv'] = str(customer['cvv'])
    if len(customer['cvv']) == 1:
        customer['cvv'] = '00' + customer['cvv']
    elif len(customer['cvv']) == 2:
        customer['cvv'] = '0' + customer['cvv']

    dummyDataPayment['buyer_name'] = customer['name']
    dummyDataPayment['credit_card_cvv'] = customer['cvv']
    dummyDataPayment['credit_card_exp'] = customer['date']
    dummyDataPayment['credit_card_number'] = customer['number']
    dummyDataPayment['payme_sale_id'] = sale.json()['payme_sale_id']

    try:
        payment = r.post(paymentUrl,data=json.dumps(dummyDataPayment),headers=headers)
        if 'status_error_details' in payment.json():
            if payment.json()['status_error_details'] == 'תוקף כרטיס אשראי לא תקין':
                new_year = new_date.year + 5
                new_date = new_date.replace(year=new_year)
                dummyDataPayment['credit_card_exp'] = new_date.strftime("%m%y")
                time.sleep(5)
                payment = r.post(paymentUrl,data=json.dumps(dummyDataPayment),headers=headers)
        counter+=1
        print(counter)
    except r.exceptions.HTTPError as e:
        print("Error on api request to pay sale:", e)
    if 'status_error_details' in payment.json():
        c['תושבת הסליקה'] = payment.json()['status_error_details']
    else:
        c['תושבת הסליקה'] = 'עבר'
    after_json.append(c)
    time.sleep(5)
print(after_json)

afterData = pd.DataFrame(after_json)
after_file_path = 'after3.xlsx'
afterData.to_excel(after_file_path, index=False)