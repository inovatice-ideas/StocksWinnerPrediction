import pandas as pd
import numpy as np
from datetime import datetime
import xlwings as xw
import input as var

temp = "templates"
stock_list = var.stock_list
ind_list = ['Volume','OBV','MACD','VWAP','ADI','NVI','PVI','PG','PVT','Volume Growth','Price Growth','Volume RSI','MFI','CMF','ATR','ADX']
ind_list.reverse()

def main(stock, cal_date):
    df = pd.read_excel(f'./{temp}/{stock}/{stock}_{cal_date.date()}.xlsx', sheet_name='MACD')
    col = list(df.columns)
    Dates = []
    Chng = []
    for i in col:
        Dates.append(i.split('(')[0])
        Chng.append(float(i.split('(')[1].split(')')[0]))
    Dates = [x for _, x in sorted(zip(Chng, Dates))]
    Chng = sorted(Chng)

    final_list = []
    for i in range(0, len(Dates)):
        for ind in ind_list:
            temp1 = []
            temp1.append(Dates[i])
            temp1.append(Chng[i])
            col_index = Dates[i] + '(' + str(Chng[i]) + ')'
            data = pd.read_excel(f'./{temp}/{stock}/{stock}_{cal_date.date()}.xlsx', sheet_name=ind)
            data = data.tail(6).reset_index(drop=True)[col_index]
            temp1.append(ind)
            temp1.extend(data)
            final_list.insert(0, temp1)
    final_temp = pd.DataFrame(final_list)
    col = {0: 'Date', 1: '% Chng', 2: 'Parameters', 3: 'Sum', 4: 'Max', 5: 'Min', 6: 'Average', 7: 'Standard Deviation', 8: 'Standard Error'}
    final_temp.rename(columns=col, inplace=True)
    final_temp.to_excel(f'{temp}/{stock}/{stock}_temp14.xlsx', index=False, freeze_panes=(1,9))

for item in stock_list:
    value = item.split("*")
    stock = value[0]
    date = datetime.strptime(value[1],'%d-%m-%Y')
    print(stock)
    print(date)
    main(stock, date)