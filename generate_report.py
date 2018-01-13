# -*- coding: utf-8 -*-

#!/usr/bin/python
# Author: Shibu Mahapatra
# Date: 10.01.2018

"""Input Date should be passed as DDMMYY"""
import urllib
import shutil
import pandas as pd
import sys
from pathlib import Path
import zipfile
import requests
import numpy as np

date = str(sys.argv[1])
fy = str(sys.argv[2])
#date='120118'
#fy='2016-17'

file_zip='EQ'+date+'_CSV.ZIP'
file='EQ'+date+'.CSV'
base_url='http://www.bseindia.com/download/BhavCopy/Equity/'+file_zip
purchase_data='purchase_data_'+fy


#delete existing files from dir
def del_tmp_files():
    for p in Path(".").glob("EQ*"):
        p.unlink()

#download file from web  
def is_downloadable(url):
    """
    Does the url contain a downloadable resource
    """
    h = requests.head(url, allow_redirects=True)
    header = h.headers
    content_type = header.get('content-type')
    if 'text' in content_type.lower():
        return False
    if 'html' in content_type.lower():
        return False
    return True


return_value = is_downloadable(base_url)

if return_value:
    request = urllib.request.Request(base_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(base_url) as response, open(file_zip, 'wb') as out_file:
        shutil.copyfileobj(response, out_file)

#unzip file
with zipfile.ZipFile(file_zip,"r") as zip_ref:
    zip_ref.extractall()

#read csv file
bhav_df = pd.read_csv(file,header=0)
bhav_filt_df = bhav_df.loc[:, ['SC_CODE', 'SC_NAME', 'CLOSE']]
#bhav_df.dtypes


purchase_df = pd.read_csv(purchase_data+'.csv', header=0)
#purchase_df.dtypes

purchase_df_n = purchase_df[purchase_df['Sell_Ind'] == 'N']
purchase_df_y = purchase_df[purchase_df['Sell_Ind'] == 'Y']


#join both pandas df
merged_df_n = pd.merge(bhav_filt_df, purchase_df_n, on='SC_CODE', how='inner')
merged_df_y = pd.merge(bhav_filt_df, purchase_df_y, on='SC_CODE', how='inner')

frames = [ merged_df_n, merged_df_y]
result = pd.concat(frames)

#filter specific columns
filt_df = result.loc[:, ['PurchaseDate' , 'SC_CODE' , 'CompanyName', 'SC_NAME' , 
                            'SharesUnits' , 'PurchasePrice' , 'Commissions' , 'CLOSE', 'Sell_Ind', 'Sold_price', 'SellDate' ]]
#rename df columns
df = filt_df.rename(columns={'SC_CODE': 'CompanyCode', 
                             'SC_NAME': 'ScriptName', 'CLOSE': 'CurrentPrice'})
    
#define CAGR
def CAGR(MarketValue, TotalCost, periods):
    try:
        cagr = (((MarketValue/TotalCost)**(1/periods)-1)*100)
    except Exception as err:
        #print ("=" * 80 + "\nSomething went wrong while calculating CAGR: {}\n".format(err) + "=" * 80)
        return 0
    else:
        return cagr

#define gain_loss_percent
def gain_loss_per(Gain_Loss, TotalCost):
    try:
        gain_loss_per = Gain_Loss / TotalCost
    except Exception as err:
        return 0
    else:
        return gain_loss_per

""" 
https://docs.scipy.org/doc/numpy-1.13.0/reference/generated/numpy.select.html
numpy.select(condlist, choicelist, default=0)[source]
"""
def marketPrice(Sell_Ind, CurrentPrice, Sold_price, SharesUnits):
    m1 = Sell_Ind=='N'
    m2 = Sell_Ind=='Y'
    a = SharesUnits * CurrentPrice
    b = SharesUnits * Sold_price
    return np.select([m1, m2], [a,b], default=0)

def durationMonths(Sell_Ind, CurrentDate, PurchaseDate, SellDate):
    m1 = Sell_Ind=='N'
    m2 = Sell_Ind=='Y'
    a = CurrentDate.to_period(freq='M') - PurchaseDate.to_period(freq='M')
    b = SellDate.to_period(freq='M') - PurchaseDate.to_period(freq='M')
    return np.select([m1, m2], [a,b], default=0)

#metrics
df['TotalCost'] = df.apply(lambda row: row.SharesUnits * row.PurchasePrice + row.Commissions, axis=1).astype(float).round(2)
df['MarketValue'] =  (df.apply(lambda row: marketPrice(row['Sell_Ind'], row['CurrentPrice'], row['Sold_price'], row['SharesUnits']), axis=1)).astype(float).round(2)
df['Gain_Loss'] = df.apply(lambda row: row.MarketValue - row.TotalCost, axis=1).round(2)
df['Gain_Loss(%)'] = df.apply(lambda row: gain_loss_per(row['Gain_Loss'], row['TotalCost']), axis=1).round(2)
df['CurrentDate'] = pd.to_datetime('today')
df['PurchaseDate'] =  pd.to_datetime(df['PurchaseDate'])
df['SellDate'] = df['SellDate'].fillna(pd.to_datetime('1900-01-01'))
df['SellDate'] = pd.to_datetime(df['SellDate'])
df['Months'] =  (df.apply(lambda row: durationMonths(row['Sell_Ind'], row['CurrentDate'], row['PurchaseDate'], row['SellDate']), axis=1))
df['CAGR'] = (df.apply(lambda row: CAGR(row['MarketValue'], row['TotalCost'], row['Months']), axis=1)).round(2)

Total_TotalCost = df['TotalCost'].sum()
Total_MarketValue = df['MarketValue'].sum()
Total_Gain_Loss = Total_MarketValue - Total_TotalCost
Total_Gain_Loss_per = Total_Gain_Loss / Total_TotalCost

final_df = df.loc[:, ['CompanyCode', 'ScriptName', 'CompanyName', 'PurchaseDate', 'SellDate', 'SharesUnits', 
                      'PurchasePrice', 'Commissions', 'CurrentPrice', 'TotalCost', 
                      'MarketValue', 'Gain_Loss', 'Gain_Loss(%)', 'CurrentDate','Months', 'CAGR']]

del_tmp_files()

# Formatting the data
from xlsxwriter.utility import xl_rowcol_to_cell

writer = pd.ExcelWriter("PortfolioAnalysis_"+fy+".xlsx", engine='xlsxwriter')
final_df.to_excel(writer, sheet_name='report', index=False)
workbook = writer.book
worksheet = writer.sheets['report']
worksheet.set_zoom(90)
# Add a number format for cells with money.
money_fmt = workbook.add_format({'num_format': '₹#,##0.00', 'bold': True})

# Add a percent format with 1 decimal point
percent_fmt = workbook.add_format({'num_format': '0.00%', 'bold': True})

#Date Format
date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd'})

# Account info columns
worksheet.set_column('B:E', 20)
# State column
worksheet.set_column('N:O', 20)

# Monthly columns
worksheet.set_column('G:L', 12, money_fmt)
# Quota percent columns
worksheet.set_column('M:M', 12, percent_fmt)

#Date Columns
worksheet.set_column('D:E', 20, date_fmt)
worksheet.set_column('N:N', 20, date_fmt)

# Define our range for the color formatting
number_rows = len(df.index)
color_range = "L2:M{}".format(number_rows+1)

# Add a format. Light red fill with dark red text.
format1 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

# Add a format. Green fill with dark green text.
format2 = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
                               
# Highlight the +ve values in Green
worksheet.conditional_format(color_range, {'type': 'cell',
                                           'criteria': '>=',
                                           'value': '0',
                                           'format': format2})
# Highlight the -ve  values in Green
worksheet.conditional_format(color_range, {'type': 'cell',
                                           'criteria': '<',
                                           'value': '0',
                                           'format': format1})

end_row=number_rows+1
worksheet.write(end_row, 8, 'TOTAL')
worksheet.write(end_row, 9, Total_TotalCost)
worksheet.write(end_row, 10, Total_MarketValue)
worksheet.write(end_row, 11, Total_Gain_Loss)
worksheet.write(end_row, 12, Total_Gain_Loss_per)
writer.save()                              