import time
import pandas as pd
from datetime import datetime
import pandas_ta as ta
import input as var
import xlwings as xw
import os

start_time = time.time()


# ========= Inputs ==================

# Items of stock_list will be in a format stock*date ... only * between stock and date
# also date will be dd-mm-yyyy ..... use - only ...

stock_list = var.stock_list
temp = "templates"
days = 14

#  ==================================



daily = var.daily
intraday = var.intraday


def datetotimestamp(date):
    time_tuple = date.timetuple()
    timestamp = round(time.mktime(time_tuple))
    return timestamp

def timestamptodate(timestamp):
    return datetime.fromtimestamp(timestamp)


def screen(stock,cal_date):
    record = []
    try:
        file = f"{daily}/{stock}.csv"
        data = pd.read_csv(file, index_col=0)
        data = data.drop(['s'], axis=1)
        data.columns = ["Datetime", "Open", "High", "Low", "Close", "Volume"]
        data['Datetime'] = data['Datetime'].apply(lambda time: timestamptodate(time).date())
        df = data.loc[(data["Datetime"] <= cal_date.date())]
        df = df.reset_index(drop=True)
        df = df[len(df)-days -1:len(df)]
        df = df.reset_index()
        for i in range(1, len(df)):
            perc = (df['Close'][i] - df['Close'][i - 1]) * 100 / df['Close'][i - 1]
            record.append({
                "Dates": df["Datetime"][i],
                "% Chng": round(perc,2)
            })
    except:
        print(f"Output No data {stock}")

    rec_df = pd.DataFrame(record)
    return rec_df


def Growth(v1,v2):
    if v1:
        return round(((v2 - v1)/v1)*100,2)
    else:
        return 0

def intraday_indicator(stock,start_date,cal_date):
    try:
        intraday_file_data = f"{intraday}/{stock}.csv"
        df = pd.read_csv(intraday_file_data, index_col=0)

        df = df.drop(['s'], axis=1)
        df.columns = ["Datetime", "Open", "High", "Low", "Close", "Volume"]
        df['Datetime'] = df['Datetime'].apply(lambda time: timestamptodate(time))
        df['Date'] = pd.to_datetime(df['Datetime']).dt.date
        df = df.loc[(df["Date"] <= cal_date.date())]

        df["PV"] = df['Volume'].shift()
        df["Volume Growth"] = df.apply(lambda x: Growth(x["PV"], x["Volume"]), axis=1)
        df = df.drop(['PV'], axis=1)

        df["PG"] = df['Close'].shift()
        df["Price Growth"] = df.apply(lambda x: Growth(x["PG"], x["Close"]), axis=1)
        df = df.drop(['PG'], axis=1)

        df["Datetime"] = pd.to_datetime(df["Datetime"])

        pg = []
        pcv = 0
        for i in range(len(df)):
            if i == 0:
                pcv = df["Close"][i]
            else:
                if df["Datetime"][i].hour == 9 and df["Datetime"][i].minute == 15:
                    pcv = df["Close"][i - 1]
            pg.append(Growth(pcv, df["Close"][i]))

        pg_df = pd.DataFrame(pg)
        df["PG"] = pg_df

        df["Volume RSI"] = ta.rsi(df["Volume"], var.vrsi_len)
        df["PVT"] = ta.pvt(df["Close"], df["Volume"])
        df["MFI"] = ta.mfi(df["High"], df["Low"], df["Close"], df["Volume"], length=var.mfi_len)
        df["CMF"] = ta.cmf(df["High"], df["Low"], df["Close"], df["Volume"], length=var.cmf_len)
        df["ADI"] = ta.ad(df["High"], df["Low"], df["Close"], df["Volume"])
        df["EOM"] = ta.eom(df["High"], df["Low"], df["Close"], df["Volume"], length=var.eom_len, divisor=var.eom_div)
        df["NVI"] = ta.nvi(df["Close"], df["Volume"], length=var.nvi_len)
        df["PVI"] = ta.pvi(df["Close"], df["Volume"], length=var.pvi_len)
        df["ATR"] = ta.atr(df["High"], df["Low"], df["Close"], length=var.atr_len)
        adx = ta.adx(df["High"], df["Low"], df["Close"], length=var.adx_len)
        df["ADX"] = adx[adx.columns[0]]

        df = df.set_index(pd.DatetimeIndex(df['Datetime']))
        df["VWAP"] = ta.vwap(df["High"], df["Low"], df["Close"], df["Volume"])

        macd = []
        macd = ta.macd(df["Close"], fast=var.macd_fast, slow=var.macd_slow, signal=var.macd_sig)
        df["MACD"] = macd[macd.columns[0]]

        df = df.loc[(df["Date"] >= start_date)]
        df["OBV"] = ta.obv(df["Close"], df["Volume"])
        df = df.reset_index(drop=True)

        return df
    except:
        pass


def indicator(df,rec_df,i_list):
    ind = pd.DataFrame()
    for i in range(len(rec_df)):
        dt = rec_df["Dates"][i]
        p = rec_df["% Chng"][i]
        data = df.loc[(df["Date"] == dt)]
        data = data.reset_index(drop=True)
        ind[f"{dt}({p})"] = data[i_list]
    return ind



def main(stock,cal_date):

    rec_df = screen(stock,cal_date)
    start_date = rec_df["Dates"][0]
    ind_list = ['Volume','Volume Growth','Price Growth','PG','OBV','Volume RSI','PVT','MFI','CMF','ADI','EOM','NVI','PVI','ATR','ADX','VWAP','MACD']
    df = intraday_indicator(stock,start_date,cal_date)

    cal_day_df = df.loc[(df["Date"] == cal_date.date())]
    cal_day_df = cal_day_df.reset_index(drop=True)

    wb = xw.Book()
    s1 = f"{start_date}_{cal_date.date()}"

    s2 = f"{stock}_{cal_date.date()}"
    wb.sheets.add(s1)
    wb.sheets.add(s2,after=s1)

    wb.sheets[s1].range('A1').options(index=False).value = df
    wb.sheets[s2].range('A1').options(index=False).value = cal_day_df

    for ind in ind_list:
        wb.sheets.add(f"{ind}",after=s2)
        ind_value = indicator(df,rec_df,ind)
        wb.sheets[ind].range('A1').options(index=False).value = ind_value

    url = temp + '/' + stock
    if not os.path.exists(url):
        os.mkdir(url)

    wb.save(f"{url}/{stock}_{cal_date.date()}.xlsx")
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