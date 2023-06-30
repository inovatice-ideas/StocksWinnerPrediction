import time
import importdata as id
import pandas as pd
from datetime import datetime
import os.path
import pandas_ta as ta
import input as var
import xlwings as xw

start_time = time.time()

stocks = var.stocks
per = var.per
d_res = var.d_res
i_res = var.i_res

daily = var.daily
intraday = var.intraday
result = var.result

daily_urls = []
int_urls = []

def datetotimestamp(date):
    time_tuple = date.timetuple()
    timestamp = round(time.mktime(time_tuple))
    return timestamp

def timestamptodate(timestamp):
    return datetime.fromtimestamp(timestamp)


# ======================================================================

def import_data():
    start = datetotimestamp(var.start_)
    end = datetotimestamp(var.end_)

    for stock in stocks:
        d_url = f"https://priceapi.moneycontrol.com/techCharts/techChartController/history?symbol={stock}&resolution={d_res}&from={start}&to={end}"
        daily_urls.append({
            "Stock": stock,
            "Url": d_url
        })
        i_url = f"https://priceapi.moneycontrol.com/techCharts/techChartController/history?symbol={stock}&resolution={i_res}&from={start}&to={end}"
        int_urls.append({
            "Stock": stock,
            "Url": i_url
        })

    key = input('Want to Import Data?(1 - Yes) : ')
    if key == '1':
        print("========= Importing Daily Data =========")
        id.import_data(daily_urls, daily)
        print("========= Importing Intraday Data =========")
        id.import_data(int_urls, intraday)

        print("--- %s seconds ---" % (time.time() - start_time))

# ===========================================================================

def screen(stock,start,end):
    record = []
    p_chg = []
    n_chg = []
    try:
        file = f"{daily}/{stock}.csv"
        data = pd.read_csv(file, index_col=0)
        data = data.drop(['s'], axis=1)
        data.columns = ["Datetime", "Open", "High", "Low", "Close", "Volume"]
        data['Datetime'] = data['Datetime'].apply(lambda time: timestamptodate(time).date())
        df = data.loc[(data["Datetime"] >= start.date()) & (data["Datetime"] <= end.date())]
        df = df.reset_index(drop=True)

        for i in range(1, len(df)):
            perc = (df['Close'][i] - df['Close'][i - 1]) * 100 / df['Close'][i - 1]
            record.append({
                "Dates": df["Datetime"][i],
                "% Chng": round(perc,2)
            })
            if perc >= var.per:
                p_chg.append({
                    "Dates": df["Datetime"][i],
                    "% Chng": round(perc, 2)
                })
            if perc <= -var.per:
                n_chg.append({
                    "Dates": df["Datetime"][i],
                    "% Chng": round(perc, 2)
                })
    except:
        print(f"Output No data {stock}")

    rec_df = pd.DataFrame(record)
    p_df = pd.DataFrame(p_chg)
    n_df = pd.DataFrame(n_chg)
    return rec_df, p_df, n_df

# ===========================================================================

def Growth(v1,v2):
    if v1:
        return round(((v2 - v1)/v1)*100,2)
    else:
        return 0

def intraday_indicator(stock):
    try:
        intraday_file_data = f"{intraday}/{stock}.csv"
        df = pd.read_csv(intraday_file_data, index_col=0)

        df = df.drop(['s'], axis=1)
        df.columns = ["Datetime", "Open", "High", "Low", "Close", "Volume"]
        df['Datetime'] = df['Datetime'].apply(lambda time: timestamptodate(time))

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

        df["OBV"] = ta.obv(df["Close"], df["Volume"])
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

        df = df.reset_index(drop=True)
        df['Date'] = pd.to_datetime(df['Datetime']).dt.date

        return df
    except:
        pass

def intraday_data(stock, start, end):

    rec_df, p_df, n_df = screen(stock, start, end)
    if not p_df.empty or not n_df.empty:
        wb = xw.Book()
        per_sheet = f"{stock}_Daily_%_Change"
        wb.sheets.add(per_sheet)
        wb.sheets[per_sheet].range('A1').options(index=False).value = rec_df

        int_df = intraday_indicator(stock)
        i_df = int_df.loc[(int_df["Date"] >= start.date()) & (int_df["Date"] <= end.date())]

        se_sheet = f"Intraday_{start.date()}_{end.date()}"
        wb.sheets.add(se_sheet, after = per_sheet)
        wb.sheets[se_sheet].range('A1').options(index=False).value = i_df

        if not n_df.empty:
            for i in range(len(n_df)):
                s_name = f'{n_df["Dates"][i]}_({n_df["% Chng"][i]})'
                data = int_df[int_df["Date"] == n_df["Dates"][i]]
                wb.sheets.add(s_name, after=se_sheet)
                wb.sheets[s_name].range('A1').options(index=False).value = data

        if not p_df.empty:
            for i in range(len(p_df)):
                s_name = f'{p_df["Dates"][i]}_({p_df["% Chng"][i]})'
                data = int_df[int_df["Date"] == p_df["Dates"][i]]
                wb.sheets.add(s_name, after=se_sheet)
                wb.sheets[s_name].range('A1').options(index=False).value = data

        wb.save(f"{result}/{stock}.xlsx")
        wb.close()


# ===========================================================================

if __name__ == "__main__":
    import_data()

    for stock in stocks:
        try:
            print(f"===== Processing {stock} ======")
            intraday_data(stock,var.s_date,var.e_date)
        except:
            pass
    print("--- %s seconds to complete---" % (time.time() - start_time))
