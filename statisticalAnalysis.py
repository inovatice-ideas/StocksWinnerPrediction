import time
import pandas as pd
from datetime import datetime
import pandas_ta as ta
import xlwings as xw
import numpy as np
import statistics
import math
import matplotlib.pyplot as plt
import templates as tp

start_time = time.time()


# ========= Inputs ==================

# Items of stock_list will be in a format stock*date ... only * between stock and date
# also date will be dd-mm-yyyy ..... use - only ...

stock_list = tp.stock_list
temp = tp.temp
days = tp.days
ind_list = ['Volume','Volume Growth','Price Growth','PG','OBV','Volume RSI','PVT','MFI','CMF','ADI','EOM','NVI','PVI','ATR','ADX','VWAP','MACD']

#  ==================================

def main(stock,cal_date):
    wb = xw.Book(f"templates/{stock}/{stock}_{cal_date.date()}.xlsx")
    df = pd.read_excel(f"templates/{stock}/{stock}_{cal_date.date()}.xlsx", sheet_name=f"{stock}_{cal_date.date()}")
    xaxis = df.index
    nvi = np.array(df['NVI'])
    pvi = np.array(df['PVI'])
    fig = plt.figure(figsize = (20, 5))
    plt.plot(xaxis, nvi, c='blue' ,label='NVI')
    plt.plot(xaxis, pvi, c='orange',label='PVI')
    plt.grid(axis='y')
    plt.legend()
    #plt.show()
    wb_temp = wb.sheets[f"{stock}_{cal_date.date()}"]
    wb_temp.pictures.add(fig, name='linechart', left=wb_temp.range('A' + str(len(df)+3)).left, top=wb_temp.range('A' + str(len(df)+3)).top, update=True)
    plt.close(fig)
    for ind in ind_list:
        df = pd.read_excel(f"templates/{stock}/{stock}_{cal_date.date()}.xlsx", sheet_name=ind)
        s = []
        max = []
        min = []
        mean = []
        stdev = []
        stderror = []
        try:
            for day in range(0, days):
                test = df[df.columns[day]]
                test.dropna(inplace=True)
                if len(test) > 1:
                    s.append(test.sum())
                    max.append(test.max())
                    min.append(test.min())
                    mean.append(test.mean())
                    stdev.append(statistics.stdev(test))
                    stderror.append((statistics.stdev(test) / math.sqrt(len(test))))
                elif len(test) == 1:
                    s.append(test.sum())
                    max.append(test.max())
                    min.append(test.min())
                    mean.append(test.mean())
                    stdev.append(0)
                    stderror.append(0)
                else:
                    s.append(0)
                    max.append(0)
                    min.append(0)
                    mean.append(0)
                    stdev.append(0)
                    stderror.append(0)
            wb.sheets[ind].range('A' + str(len(df)+3)).options(index=False).value = s
            wb.sheets[ind].range('A' + str(len(df)+4)).options(index=False).value = max
            wb.sheets[ind].range('A' + str(len(df)+5)).options(index=False).value = min
            wb.sheets[ind].range('A' + str(len(df)+6)).options(index=False).value = mean
            wb.sheets[ind].range('A' + str(len(df)+7)).options(index=False).value = stdev
            wb.sheets[ind].range('A' + str(len(df)+8)).options(index=False).value = stderror
            xaxis = np.array(df.columns)
            fig = plt.figure()
            plt.bar(xaxis, stdev)
            plt.title("STDEV")
            fig.autofmt_xdate()
            plt.grid(axis='y')
            #plt.show()
            wb.sheets[ind].pictures.add(fig, name='Stdev', left=wb.sheets[ind].range('A' + str(len(df)+10)).left, top=wb.sheets[ind].range('A' + str(len(df)+10)).top, update=True)
            plt.close(fig)
            fig = plt.figure()
            plt.bar(xaxis, stderror)
            plt.title("STDERROR")
            fig.autofmt_xdate()
            plt.grid(axis='y')
            #plt.show()
            wb.sheets[ind].pictures.add(fig, name='Stderror', left=wb.sheets[ind].range('K' + str(len(df)+10)).left, top=wb.sheets[ind].range('K' + str(len(df)+10)).top, update=True)
            plt.close(fig)
            wb.save(f"templates/{stock}/{stock}_{cal_date.date()}.xlsx")
        except:
            pass
    wb.close()

# for stock in stocks:
#     main(stock,c_date)

for item in stock_list:
    value = item.split("*")
    stock = value[0]
    date = datetime.strptime(value[1],'%d-%m-%Y')
    print(stock)
    print(date)
    main(stock, date)