# This script automatically extracts stock market data from the MarketWatch website.
import os
import re
import time
import urllib.request
import pandas as pd
#
os.chdir('C:/Users/Michael/Dropbox/Documents/Financial/Investments')
symbols_in = pd.read_excel('./symbols_in.xlsx')
#
df = pd.DataFrame(columns=['symbol','name','industry','sector','dividend','price','target','buy','overweight','hold','underweight','sell'])
df['symbol'] = symbols_in['symbol']
df.set_index('symbol', inplace=True)

symbols = symbols_in['symbol'].to_list()
for symbol in symbols:
    print(symbol)
    time.sleep(0.5)  # pause briefly between stock symbols to respect server traffic
    
    # read dividend info from stock homepage
    try:
        fhand = urllib.request.urlopen('https://www.marketwatch.com/investing/stock/' + symbol)
        stage = 0
        for line in fhand:
            linetext = line.decode().strip()
            if stage == 'dividend':
                try:
                    df.loc[symbol,'dividend'] = float(re.findall('<span class="primary ">\$(.+)</span>', linetext)[0])
                except:
                    df.loc[symbol,'dividend'] = float(0)
                break
            if '<small class="label">Dividend</small>' in linetext:
                stage = 'dividend'
    except:
        continue
    
    # read company profile page
    try:
        fhand = urllib.request.urlopen('https://www.marketwatch.com/investing/stock/' + symbol + '/company-profile')
        stage = 0
        for line in fhand:
            linetext = line.decode().strip()
            # find industry and sector
            if stage == 'industry':
                df.loc[symbol,'industry'] = re.findall('<span class="primary ">(.+)</span>', linetext)[0]
                stage = 0
            elif stage == 'sector':
                df.loc[symbol,'sector'] = re.findall('<span class="primary ">(.+)</span>', linetext)[0]
                break  # all done here; move along to next URL
            if '<small class="label">Industry</small>' in linetext:
                stage = 'industry'
            elif '<small class="label">Sector</small>' in linetext:
                stage = 'sector'
    except:
        continue
    
    # read analyst estimates page
    try:
        fhand = urllib.request.urlopen('https://www.marketwatch.com/investing/stock/' + symbol + '/analystestimates')
        stage = 'name';  toggle = False;  count = 0
        for line in fhand:
            linetext = line.decode().strip()

            # find company name
            if stage == 'name':
                if '<h1 class="company__name">' in linetext:
                    df.loc[symbol,'name'] = re.findall('<h1 class="company__name">(.+)</h1>', linetext)[0].replace('amp;', '')
                    stage = 'price'

            # find latest price
            elif stage == 'price':
                if '<bg-quote class="value"' in linetext:
                    df.loc[symbol,'price'] = float(re.findall('.+>(.+)</bg-quote>', linetext)[0].replace(',', ''))
                    stage = 'target'
                elif '<span class="value"' in linetext:
                    df.loc[symbol,'price'] = float(re.findall('.+>(.+)</span>', linetext)[0].replace(',', ''))
                    stage = 'target'
            
            # find average price target
            elif stage == 'target':
                if toggle:
                    df.loc[symbol,'target'] = float(re.findall('.+>(.+)</td>', linetext)[0].replace(',', ''))
                    stage = 'buy'
                elif 'Average Target Price' in linetext:
                    toggle = True

            # find number of 'Buy' recommendations
            elif stage == 'buy':
                if '<td class="table__cell w25">Buy</td>' in linetext:
                    count = 0
                elif '<span class="value' in linetext:
                    count = count + 1
                if count == 3:
                    count = 0
                    df.loc[symbol,'buy'] = pd.to_numeric(re.findall('.+>(.+)</span>', linetext)[0])
                    stage = 'over'
            
            # find number of 'Overweight' recommendations
            elif stage == 'over':
                if '<td class="table__cell w25">Overweight</td>' in linetext:
                    count = 0
                elif '<span class="value' in linetext:
                    count = count + 1
                if count == 3:
                    count = 0
                    df.loc[symbol,'overweight'] = pd.to_numeric(re.findall('.+>(.+)</span>', linetext)[0])
                    stage = 'hold'
            
            # find number of 'Hold' recommendations
            elif stage == 'hold':
                if '<td class="table__cell w25">Hold</td>' in linetext:
                    count = 0
                elif '<span class="value' in linetext:
                    count = count + 1
                if count == 3:
                    count = 0
                    df.loc[symbol,'hold'] = pd.to_numeric(re.findall('.+>(.+)</span>', linetext)[0])
                    stage = 'under'
            
            # find number of 'Underweight' recommendations
            elif stage == 'under':
                if '<td class="table__cell w25">Underweight</td>' in linetext:
                    count = 0
                elif '<span class="value' in linetext:
                    count = count + 1
                if count == 3:
                    count = 0
                    df.loc[symbol,'underweight'] = pd.to_numeric(re.findall('.+>(.+)</span>', linetext)[0])
                    stage = 'sell'
            
            # find number of 'Sell' recommendations
            elif stage == 'sell':
                if '<td class="table__cell w25">Sell</td>' in linetext:
                    count = 0
                elif '<span class="value' in linetext:
                    count = count + 1
                if count == 3:
                    count = 0
                    df.loc[symbol,'sell'] = pd.to_numeric(re.findall('.+>(.+)</span>', linetext)[0])
                    break  # all done here; move along to next stock symbol
    except:
        continue

df['pct_gain'] = 100*(df['target'] - df['price']) / df['price']
df['buy_pct'] = 100*df['buy'] / (df['buy'] + df['overweight'] + df['hold'] + df['underweight'] + df['sell'])
df['sell_pct'] = 100*df['sell'] / (df['buy'] + df['overweight'] + df['hold'] + df['underweight'] + df['sell'])
df = df.sort_values(['buy_pct','pct_gain'], ascending=False)

df.to_excel('./output.xlsx', float_format="%.2f")
print('Data scraping successfully completed.')
