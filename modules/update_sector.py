import os
from yahoofinancials import YahooFinancials
import yahoo_fin.stock_info as si
import yfinance as yf
import pandas as pd
import numpy as np
from tqdm import tqdm
import math
import time 
from openpyxl import load_workbook
from datetime import datetime

path = 'myrics_main_light3_processed.xlsx'

df = pd.read_excel(path, index_col=0, sheet_name='Main')

i = 0

#df = df[df['Cheap Asset']==True]
#tickers = list(df[df['Industry'].isnull()].index)
tickers = list(df[df['Industry']=="#bad"].index)

for tick in tqdm(tickers[0:1000]):
    
    try:
        #df.loc[tick, 'Category Name'] = yf.Ticker(tick).info['sector']
        insinfo = yf.Ticker(tick).info
        df.loc[tick, 'Industry'] = insinfo['industry']
        df.loc[tick, 'Name2'] = insinfo['longName']
        # 1E4.SG
    except:
        df.loc[tick, 'Industry'] = "#bad"
    
    #df.loc[tick, 'Updated'] = datetime.now().date()

    i += 1
    if i%200==0 or i>=len(tickers):
        df.to_csv('main_draft_light.csv', header=True, index_label=True)
    
    insinfo = {} # clear dict from the previous instrument info
    time.sleep(0.6)


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
