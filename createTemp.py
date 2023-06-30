import pandas as pd
import numpy as np
from datetime import datetime
import xlwings as xw
import input as var
import os

temp = "templates"
stocks = var.stocks
ind_list = ['Volume','OBV','MACD','VWAP','ADI','NVI','PVI','PG','PVT','Volume Growth','Price Growth','Volume RSI','MFI','CMF','ATR','ADX']

df = pd.read_excel(f'{temp}/DailyBreakout.xlsx')

def main(stock):
    df1 = df.groupby(['Stock']).get_group(stock)
    df1.sort_values(by = '% Chng', ascending=False, inplace=True)
    index_order = df1.index
    final_list = []
    for i in index_order:
        for ind in ind_list:
            temp1 = []
            temp1.append(df1['Date'][i].date())
            temp1.append(df1['% Chng'][i])
            col_index = str(df1['Date'][i].date()) + '(' + str(df1['% Chng'][i]) + ')'
            data = pd.read_excel(f'./{temp}/{stock}/{stock}.xlsx', sheet_name=ind)
            data = data.tail(6).reset_index(drop=True)[col_index]
            temp1.append(ind)
            temp1.extend(data)
            final_list.append(temp1)
    final_temp = pd.DataFrame(final_list)
    col = {0: 'Date', 1: '% Chng', 2: 'Parameters', 3: 'Sum', 4: 'Max', 5: 'Min', 6: 'Average', 7: 'Standard Deviation', 8: 'Standard Error'}
    final_temp.rename(columns=col, inplace=True)
    final_temp.to_excel(f'{temp}/{stock}/{stock}_temp.xlsx', index=False, freeze_panes=(1,9))

for item in stocks:
    print(item)
    main(item)