import aiohttp , asyncio
import pandas as pd
import sys

async def get_data(session, url,stock,folder):
    async with session.get(url) as resp:
        try:
            data = await resp.json()
            pdata = pd.DataFrame(data)
            pdata.to_csv(f'{folder}/{stock}.csv')
            print(f'Importing {stock}')
        except:
            print(f"Problem  {stock}")

async def main(stock_urls,folder):
    connector = aiohttp.TCPConnector(limit=10)
    async with aiohttp.ClientSession(connector=connector) as session:
        tasks = []
        for i in range(0,len(stock_urls)):
            # url = f"https://priceapi.moneycontrol.com/techCharts/techChartController/history?symbol={stock}&resolution={res}&from={start}&to={end}"
            try:
                tasks.append(asyncio.ensure_future(get_data(session, stock_urls[i]['Url'],stock_urls[i]['Stock'],folder)))
            except:
                print(f"Main Problem {stock}")
        out = await asyncio.gather(*tasks)
        return out


if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
else:
    asyncio.SelectorEventLoop()


def import_data(stock_urls,folder):
    asyncio.run(main(stock_urls,folder))

