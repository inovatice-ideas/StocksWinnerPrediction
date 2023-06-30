import pandas as pd
from datetime import datetime
import input as var
import xlwings as xw

# ======== Inputs ===============================================

per_range = [3, 22]  # It means 5 to 10 and also -10 to -5
start = datetime(2022, 1, 1)
end = datetime(2022, 12, 9)
temp = "templates"
# ===============================================================

stocks = var.stocks
daily = var.daily

record = []


def timestamptodate(timestamp):
    return datetime.fromtimestamp(timestamp).date()


fd = pd.DataFrame()


def slice_data(stock):
    file = f"{daily}/{stock}.csv"
    data = pd.read_csv(file, index_col=0)
    data['t'] = data['t'].apply(lambda time: timestamptodate(time))
    df = data.loc[(data["t"] <= end.date())]
    df = df.reset_index(drop=True)

    row = []
    for i in range(1, len(df)):
        perc = round((df['c'][i] - df['c'][i - 1]) * 100 / df['c'][i - 1], 2)
        if abs(perc) > per_range[0] and abs(perc) < per_range[1]:
            row.append({
                "Stock" : stock,
                "% Chng" : perc,
                "Date" : df['t'][i]
            })

    out = pd.DataFrame(row)
    if not out.empty:
        out = out.loc[(out["Date"] >= start.date())]
        out = out.reset_index(drop=True)

    return out


record = pd.DataFrame()

for stock in stocks:
    data = slice_data(stock)
    if not data.empty:
        record = pd.concat([record,data])
record = record.reset_index(drop=True)

wb = xw.Book()
wb.sheets[0].range('A1').options(index=False).value = record
wb.save(f"{temp}/DailyBreakout.xlsx")
wb.close()
