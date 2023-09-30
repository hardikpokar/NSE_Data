import requests
import pandas as pd
import time
import matplotlib.pyplot as plt
import openpyxl
import schedule as schedule
import xlsxwriter as xlsxwriter
import xlwings as xw
import xlsxwriter
from xlsxwriter import workbook
from openpyxl import load_workbook

url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (HTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    'accept-language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
    'accept-encoding': 'gzip, deflate, br'}

session = requests.Session()

def importdata():
    request = session.get(url, headers=headers)
    cookies = dict(request.cookies)
    response = session.get(url, headers=headers, cookies=cookies).json()
    rdata = pd.DataFrame(response)
    t_ce = rdata['filtered']['CE']['totOI']
    t_pe = rdata['filtered']['PE']['totOI']
    value1 = rdata["records"]["underlyingValue"]
    value = value1
    dt = rdata['records']['timestamp']
    trend = t_pe - t_ce
    calltoput = (t_pe / t_ce )
    t = dt.split(" ")
    value_new = value
    data = {
        "Time": t[1],
        "COI": t_ce,
        "POI": t_pe,
        "Trend": trend,
        "calltoput" : calltoput,
        "value_new" : value_new,
    }
    return data

print("|---------------------------------------------------------------------------|")
print("|{:<9}| {:<12}| {:<15} | {:<15}| {:<12} | {:<15}|".format(" Time", "underlying_value", " Total Call OI", " Total Put OI", "Trend" , "calltoputratio" ))

print("|---------------------------------------------------------------------------|")
while True:
        data = importdata()
        print("|{:<9}| {:<12} |  {:<12}|    {:<12}| {:<10} | {:<10}|".format(data["Time"],data["value_new"], data["COI"], data["POI"], data["Trend"], data["calltoput"]))
        print("|---------------------------------------------------------------------------|")
        writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')
        writer.close()
        data = importdata()
        df = pd.DataFrame({'Time': [data["Time"]],
                           'Underlying_Value': [data["value_new"]],
                           'Total_CE': [data["COI"]],
                           'Total_PE': [data["POI"]],
                           'TREND': [data["Trend"]],
                           'PUT_CALL_RATIO': [data["calltoput"]]})
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1', index=True)

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        time.sleep(10)
