import os
from yahoofinancials import YahooFinancials
import yahoo_fin as yf
import yahoo_fin.stock_info as si
import pandas as pd
import numpy as np
from tqdm import tqdm
import math
import time
from openpyxl import load_workbook
from datetime import datetime

multiplier = {"K": 1000,
                  "M": 1000000,
                  "B": 1000000000,
                  "T": 1000000000000}

path = 'myrics_main.xlsx'

df = pd.read_excel(path, index_col=0, sheet_name='Main')

i = 0
#for tick in tqdm(list(df[df['Portfolio']==1].index)): ['HFBL', 'MRAAF', '0QHK.L', 'PEKB.F', 'TQY.F', '1KR.F', 'SLFPY', 'HER.MI', 'DIP.F', '2763.T']
#tickers = list(df[df['Interesting'].notnull()].index)
#tickers = list(df[df['Comment'].notnull()].index)
tickers = list(df[df['Update Status'].isnull()].index) 
#tickers.extend(list(df[df['Comment'].notnull()][df[df['Comment'].notnull()]['Update Status'].isnull()].index))
for tick in tqdm(tickers):
    
    try:
        temp = si.get_stats(tick)
        #temp = si.get_stats('21P1.F')

        # Spot
        try:
            quote = si.get_quote_table(tick)
            price = float(quote["Quote Price"])
            df.loc[tick, 'Spot'] = price
        except:
            df.loc[tick, 'Spot'] = ""

        #currency

        df.loc[tick, 'ROA'] = float(str(temp[temp.Attribute.str.contains("Return on Assets")]['Value'].to_numpy()[0]).replace("%", ""))
            
        df.loc[tick, 'ROE'] = float(str(temp[temp.Attribute.str.contains("Return on Equity")]['Value'].to_numpy()[0]).replace("%", ""))
        
        df.loc[tick, 'P/E Ratio'] = float(quote["PE Ratio (TTM)"])
        
        cap = float(str(quote["Market Cap"]).replace("M", "").replace("B", "").replace("T", ""))
        df.loc[tick, 'Market Cap'] = cap * multiplier[str(quote["Market Cap"])[-1]]

        try:
            low_52 = float(str(quote["52 Week Range"]).split("-")[0])
            df.loc[tick, '52 Range low'] = low_52
        except:
            df.loc[tick, '52 Range low'] = "NaN"
            
        try:
            high_52 = float(str(quote["52 Week Range"]).split("-")[1])
            df.loc[tick, '52 Range high'] = high_52
        except:
            df.loc[tick, '52 Range high']= "NaN"

        try:
            df.loc[tick, 'Spot_Ratio'] = price/((low_52  + high_52)/2)
        except:
            df.loc[tick, 'Spot_Ratio'] = "NaN"

        ################################################################
        try:
            df.loc[tick, '200-Day Moving Average'] = float(str(temp[temp.Attribute.str.contains("200-Day Moving Average")]['Value'].to_numpy()[0]))
        except:
            continue

        try:
            df.loc[tick, '% Held by Institutions 1'] = float(str(temp[temp.Attribute.str.contains("% Held by Institutions 1")]['Value'].to_numpy()[0]).replace("%", ""))
        except:
            continue

        try:
            revenue = float(str(temp[temp.Attribute.str.contains("Revenue")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Revenue'] = revenue * multiplier[str(temp[temp.Attribute.str.contains("Revenue")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            ebitda = float(str(temp[temp.Attribute.str.contains("EBITDA")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'EBITDA'] = ebitda * multiplier[str(temp[temp.Attribute.str.contains("EBITDA")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            df.loc[tick, 'Total Debt/Equity'] = float(str(temp[temp.Attribute.str.contains("Total Debt/Equity")]['Value'].to_numpy()[0]))
            df.loc[tick, 'Current Ratio'] = float(str(temp[temp.Attribute.str.contains("Current Ratio")]['Value'].to_numpy()[0]))
        except:
            continue

        try:
            df.loc[tick, '% Quarterly Revenue Growth'] = float(str(temp[temp.Attribute.str.contains("Quarterly Revenue Growth")]['Value'].to_numpy()[0]).replace("%", ""))
        except:
            continue

        try:
            gross_profit = float(str(temp[temp.Attribute.str.contains("Gross Profit")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Gross Profit'] = gross_profit * multiplier[str(temp[temp.Attribute.str.contains("Gross Profit")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            Total_Cash = float(str(temp[temp.Attribute.str.contains("Total Cash")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Total Cash'] = Total_Cash * multiplier[str(temp[temp.Attribute.str.contains("Total Cash")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            Total_Debt = float(str(temp[temp.Attribute.str.contains("Total Debt")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Total Debt'] = Total_Debt * multiplier[str(temp[temp.Attribute.str.contains("Total Debt")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            df.loc[tick, 'Book Value Per Share'] = float(str(temp[temp.Attribute.str.contains("Book Value Per Share")]['Value'].to_numpy()[0]))
        except:
            continue

        try:
            op_cashflow = float(str(temp[temp.Attribute.str.contains("Operating Cash Flow")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Operating Cash Flow'] = op_cashflow * multiplier[str(temp[temp.Attribute.str.contains("Operating Cash Flow")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        try:
            lev_cashflow = float(str(temp[temp.Attribute.str.contains("Levered Free Cash Flow")]['Value'].to_numpy()[0]).replace("M", "").replace("B", "").replace("T", ""))
            df.loc[tick, 'Levered Free Cash Flow'] = lev_cashflow * multiplier[str(temp[temp.Attribute.str.contains("Levered Free Cash Flow")]['Value'].to_numpy()[0])[-1]]
        except:
            continue

        ################################################################
        df.loc[tick, 'Update Status'] = "#good2"
    except:
        df.loc[tick, 'Update Status'] = "#bad2"
    
    df.loc[tick, 'Updated'] = datetime.now().date()

    # total assets and total liabilities
    try:
        bs = si.get_balance_sheet(tick)
        df.loc[tick, 'Total Assets'] = bs.loc["totalAssets"][0]
        df.loc[tick, 'Total Liabilities'] = bs.loc["totalLiab"][0]
        # marcet cap < (total assets - total liabilities) * 1.5 #cheap assets
        df.loc[tick, 'Cheap Assets (1.5(TA-TL))'] = 1.5 * (bs.loc["totalAssets"][0] - bs.loc["totalLiab"][0])
    except:
        df.loc[tick, 'Update Status'] = "#bs_bad2"

    # to calc earnings growth
    # ebit = si.get_income_statement(tick).loc['ebit'][]
    # for i in len(ebit)-1:
    #   eb = si.get_income_statement(tick).loc['ebit'][i]
    #   eb1 = si.get_income_statement(tick).loc['ebit'][i+1] 
    #   eb_growth = (eb1/eb)-1

    i += 1
    if i%200==0 or i>=len(tickers):
        df.to_csv('main_draft.csv', header=True, index_label=True)

    time.sleep(1)


#df_pf = df[df['Comment'].notnull()]
#df_pf = df[df['Portfolio']==1]

# write to workbook
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book

# delet existing "Main" sheet
std = book['Main']
book.remove(std)

# write df to a new Main sheet
df.to_excel(writer, sheet_name='Main')
#df_pf.to_excel(writer, sheet_name='pf')
writer.save()
writer.close()
