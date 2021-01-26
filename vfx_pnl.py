#!/usr/bin/env python
# coding: utf-8

# In[1]:


from datetime import datetime
from datetime import timedelta
import pandas as pd
from sqlalchemy import create_engine
from xlsx2csv import Xlsx2csv
import os 
import xlwings as xw
from xlwings.constants import AutoFillType
from pathlib import Path
from xlwings.constants import DeleteShiftDirection
from xlwings.constants import Direction
import time
from xlwings.constants import InsertShiftDirection

if datetime.today().strftime("%A")=='Monday' :
    DD=(datetime.today()-timedelta(days=3)).date()
    DDY=(datetime.today()-timedelta(days=4)).date()
    DDN=(datetime.today()).date()
elif datetime.today().strftime("%A")=='Tuesday' :
    DD=(datetime.today()-timedelta(days=1)).date()
    DDY=(datetime.today()-timedelta(days=4)).date()
    DDN=(datetime.today()).date()
else :
    DD=(datetime.today()-timedelta(days=1)).date()
    DDY=(datetime.today()-timedelta(days=2)).date()
    DDN=(datetime.today()).date()

NY=DDN.strftime("%Y")
NM=DDN.strftime("%m")
ND=DDN.strftime("%d")
TY=DD.strftime("%Y")
TM=DD.strftime("%m")
TD=DD.strftime("%d")
TyY=DDY.strftime("%Y")
TyM=DDY.strftime("%m")
TyD=DDY.strftime("%d")

# NY='2020'
# NM='09'
# ND='11'
# TY='2020'
# TM='09'
# TD='10'
# TyY='2020'
# TyM='09'
# TyD='09'


# In[8]:


def xlsx2csv(excel_path,csv_path,sheet_name):
    sid=Xlsx2csv(excel_path).getSheetIdByName(name=sheet_name)
    Xlsx2csv(excel_path).convert(csv_path,sheetid=sid)
    df = pd.read_csv(csv_path,header=None).dropna(how='all')
    os.remove(csv_path)
    return df
vfx_df = xlsx2csv(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\report\For_PNL-{}{}.xlsx'                  .format(TY,TM,TM,TD,TM,TD),'vfx.csv','VFX')
vfx2_df = xlsx2csv(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\report\For_PNL-{}{}.xlsx'                  .format(TY,TM,TM,TD,TM,TD),'vfx2.csv','VFX2')
res_df = xlsx2csv(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\report\For_PNL-{}{}.xlsx'                  .format(TY,TM,TM,TD,TM,TD),'res.csv','Result')
mt5_df = xlsx2csv(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\report\For_PNL-{}{}.xlsx'                  .format(TY,TM,TM,TD,TM,TD),'mt5.csv','MT5')
raw_df = xlsx2csv(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\report\For_PNL-{}{}.xlsx'                  .format(TY,TM,TM,TD,TM,TD),'raw.csv','Raw')


# In[9]:


vfx1=[]
vfx2=[]
mal1=[]
mal2=[]
vtl1=[]
vtl2=[]
s_c=[]
m_c=[]
m_c1=[]
mt5=[]
res=[]



vfx1.append(None)
vfx1.append(float(vfx_df.iloc[2,2]))
vfx1.append(0)
vfx1.append(float(res_df.iloc[1,7]))
vfx1.append(None)
vfx1.append(0)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(vfx_df.iloc[4,2]))
vfx1.append(float(vfx_df.iloc[5,2]))
vfx1.append(float(res_df.iloc[2,7]))
vfx1.append(float(res_df.iloc[3,7]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(vfx_df.iloc[7,2]))
vfx1.append(float(vfx_df.iloc[8,2]))
vfx1.append(float(res_df.iloc[4,7]))
vfx1.append(float(res_df.iloc[5,7]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(vfx_df.iloc[10,2]))
vfx1.append(float(vfx_df.iloc[11,2]))
vfx1.append(float(vfx_df.iloc[12,2]))
vfx1.append(float(vfx_df.iloc[13,2]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(res_df.iloc[6,7]))
vfx1.append(float(res_df.iloc[7,7]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(vfx_df.iloc[15,2]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(res_df.iloc[8,7]))
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(None)
vfx1.append(float(res_df.iloc[9,7]))
vfx1.append(float(res_df.iloc[10,7]))


vfx2.append(None)
vfx2.append(float(vfx2_df.iloc[2,2]))
vfx2.append(0)
vfx2.append(float(res_df.iloc[1,8]))
vfx2.append(None)
vfx2.append(0)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(vfx2_df.iloc[3,2]))
vfx2.append(float(vfx2_df.iloc[4,2]))
vfx2.append(float(res_df.iloc[2,8]))
vfx2.append(float(res_df.iloc[3,8]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(vfx2_df.iloc[5,2]))
vfx2.append(float(vfx2_df.iloc[6,2]))
vfx2.append(float(res_df.iloc[4,8]))
vfx2.append(float(res_df.iloc[5,8]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(vfx2_df.iloc[7,2]))
vfx2.append(float(vfx2_df.iloc[8,2]))
vfx2.append(float(vfx2_df.iloc[9,2]))
vfx2.append(float(vfx2_df.iloc[10,2]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(res_df.iloc[6,8]))
vfx2.append(float(res_df.iloc[7,8]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(vfx2_df.iloc[11,2]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(res_df.iloc[8,8]))
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(None)
vfx2.append(float(res_df.iloc[9,8]))
vfx2.append(float(res_df.iloc[10,8]))


mal1.append(None)
mal1.append(float(vfx_df.iloc[2,10]))
mal1.append(0)
mal1.append(float(res_df.iloc[1,9]))
mal1.append(None)
mal1.append(0)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(vfx_df.iloc[4,10]))
mal1.append(float(vfx_df.iloc[5,10]))
mal1.append(float(res_df.iloc[2,9]))
mal1.append(float(res_df.iloc[3,9]))
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(vfx_df.iloc[7,10]))
mal1.append(float(vfx_df.iloc[8,10]))
mal1.append(float(res_df.iloc[4,9]))
mal1.append(float(res_df.iloc[5,9]))
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(vfx_df.iloc[10,10]))
mal1.append(float(vfx_df.iloc[11,10]))
mal1.append(float(vfx_df.iloc[12,10]))
mal1.append(float(vfx_df.iloc[13,10]))
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(res_df.iloc[6,9]))
mal1.append(float(res_df.iloc[7,9]))
mal1.append(None)
mal1.append(None)
mal1.append(float(vfx_df.iloc[15,10]))
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(res_df.iloc[8,9]))
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(None)
mal1.append(float(res_df.iloc[9,9]))
mal1.append(float(res_df.iloc[10,9]))


mal2.append(None)
mal2.append(float(vfx2_df.iloc[2,10]))
mal2.append(0)
mal2.append(float(res_df.iloc[1,11]))
mal2.append(None)
mal2.append(0)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(vfx2_df.iloc[3,10]))
mal2.append(float(vfx2_df.iloc[4,10]))
mal2.append(float(res_df.iloc[2,11]))
mal2.append(float(res_df.iloc[3,11]))
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(vfx2_df.iloc[5,10]))
mal2.append(float(vfx2_df.iloc[6,10]))
mal2.append(float(res_df.iloc[4,11]))
mal2.append(float(res_df.iloc[5,11]))
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(vfx2_df.iloc[7,10]))
mal2.append(float(vfx2_df.iloc[8,10]))
mal2.append(float(vfx2_df.iloc[9,10]))
mal2.append(float(vfx2_df.iloc[10,10]))
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(res_df.iloc[6,11]))
mal2.append(float(res_df.iloc[7,11]))
mal2.append(None)
mal2.append(None)
mal2.append(float(vfx2_df.iloc[11,10]))
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(res_df.iloc[8,11]))
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(None)
mal2.append(float(res_df.iloc[9,11]))
mal2.append(float(res_df.iloc[10,11]))


vtl1.append(None)
vtl1.append(float(vfx_df.iloc[2,18]))
vtl1.append(0)
vtl1.append(float(res_df.iloc[1,10]))
vtl1.append(None)
vtl1.append(0)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(vfx_df.iloc[4,18]))
vtl1.append(float(vfx_df.iloc[5,18]))
vtl1.append(float(res_df.iloc[2,10]))
vtl1.append(float(res_df.iloc[3,10]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(vfx_df.iloc[7,18]))
vtl1.append(float(vfx_df.iloc[8,18]))
vtl1.append(float(res_df.iloc[4,10]))
vtl1.append(float(res_df.iloc[5,10]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(vfx_df.iloc[10,18]))
vtl1.append(float(vfx_df.iloc[11,18]))
vtl1.append(float(vfx_df.iloc[12,18]))
vtl1.append(float(vfx_df.iloc[13,18]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(res_df.iloc[6,10]))
vtl1.append(float(res_df.iloc[7,10]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(vfx_df.iloc[15,18]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)
vtl1.append(float(res_df.iloc[8,10]))
vtl1.append(None)
vtl1.append(None)
vtl1.append(None)


vtl2.append(None)
vtl2.append(float(vfx2_df.iloc[2,18]))
vtl2.append(0)
vtl2.append(float(res_df.iloc[1,12]))
vtl2.append(None)
vtl2.append(0)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(vfx2_df.iloc[3,18]))
vtl2.append(float(vfx2_df.iloc[4,18]))
vtl2.append(float(res_df.iloc[2,12]))
vtl2.append(float(res_df.iloc[3,12]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(vfx2_df.iloc[5,18]))
vtl2.append(float(vfx2_df.iloc[6,18]))
vtl2.append(float(res_df.iloc[4,12]))
vtl2.append(float(res_df.iloc[5,12]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(vfx2_df.iloc[7,18]))
vtl2.append(float(vfx2_df.iloc[8,18]))
vtl2.append(float(vfx2_df.iloc[9,18]))
vtl2.append(float(vfx2_df.iloc[10,18]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(res_df.iloc[6,12]))
vtl2.append(float(res_df.iloc[7,12]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(vfx2_df.iloc[11,18]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)
vtl2.append(float(res_df.iloc[8,12]))
vtl2.append(None)
vtl2.append(None)
vtl2.append(None)


s_c.append(float(vfx_df.iloc[36,18]))
s_c.append(None)
s_c.append(float(vfx_df.iloc[36,12]))
s_c.append(float(vfx_df.iloc[49,18]))
s_c.append(None)

m_c.append(float(mt5_df.iloc[5,5]))
m_c.append(None)
m_c.append(float(mt5_df.iloc[5,9]))
m_c.append(float(mt5_df.iloc[7,5]))
m_c.append(None)

m_c1.append(float(mt5_df.iloc[2,0]))
m_c1.append(float(mt5_df.iloc[2,1]))
m_c1.append(float(mt5_df.iloc[6,0]))
m_c1.append(float(mt5_df.iloc[6,1]))


mt5.append(None)
mt5.append(float(mt5_df.iloc[5,5]))
mt5.append(0)
mt5.append(float(raw_df.iloc[6,14]))
mt5.append(None)
mt5.append(0)
mt5.append(None)
mt5.append(None)
mt5.append(float(mt5_df.iloc[5,7]))
mt5.append(float(mt5_df.iloc[7,7]))
mt5.append(0)
mt5.append(0)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(float(mt5_df.iloc[5,6]))
mt5.append(float(mt5_df.iloc[7,6]))
mt5.append(0)
mt5.append(0)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(0)
mt5.append(None)
mt5.append(0)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(0)
mt5.append(0)
mt5.append(None)
mt5.append(None)
mt5.append(float(mt5_df.iloc[7,5]))
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(None)
mt5.append(float(raw_df.iloc[6,13]))


res.append(float(res_df.iloc[11,7]))
res.append(float(res_df.iloc[12,7]))
res.append(float(res_df.iloc[13,7]))
res.append(float(res_df.iloc[14,7]))
res.append(float(res_df.iloc[11,9]))
res.append(float(res_df.iloc[12,9]))
res.append(float(res_df.iloc[13,9]))
res.append(float(res_df.iloc[14,9]))
res.append(float(res_df.iloc[11,10]))
res.append(float(res_df.iloc[12,10]))
res.append(float(res_df.iloc[13,10]))
res.append(float(res_df.iloc[14,10]))
res.append(float(res_df.iloc[11,8]))
res.append(float(res_df.iloc[12,8]))
res.append(float(res_df.iloc[13,8]))
res.append(float(res_df.iloc[14,8]))
res.append(float(res_df.iloc[11,11]))
res.append(float(res_df.iloc[12,11]))
res.append(float(res_df.iloc[13,11]))
res.append(float(res_df.iloc[14,11]))
res.append(float(res_df.iloc[11,12]))
res.append(float(res_df.iloc[12,12]))
res.append(float(res_df.iloc[13,12]))
res.append(float(res_df.iloc[14,12]))


# In[10]:


# pd.set_option('display.max_rows', 60)
# pd.set_option('display.max_columns', 60)


# In[11]:


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

app = xw.App(visible=False)
wb = app.books.open(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\PNL\VFX-PNL-{}{}{}.xlsx'.format(TyY,TyM,TyM,TyD,TyY,TyM,TyD))

s_c_sh=wb.sheets['核對-S']
m_c_sh=wb.sheets['核對-mt5']
res_sh=wb.sheets['result']
vfx_sh=wb.sheets['匯總']
vfx2_sh=wb.sheets['匯總2']
mal_sh=wb.sheets['匯總 -MAL']
vtl_sh=wb.sheets['匯總 -VTL']
mal2_sh=wb.sheets['匯總 -MAL 2']
vtl2_sh=wb.sheets['匯總 -VTL 2']
mt5_sh=wb.sheets['匯總-mt5']


max_col1=res_sh.cells(2, "C").end('right').column
res_sh.api.columns(max_col1+1).insert

max_col2=res_sh.cells(2, f"{colnum_string(max_col1+6)}").end('right').column
res_sh.api.columns(max_col2+1).insert

max_col3=res_sh.cells(2, f"{colnum_string(max_col2+6)}").end('right').column
res_sh.api.columns(max_col3+1).insert

max_col4=vfx_sh.cells(3, "F").end('right').column
vfx_sh.api.columns(max_col4+1).insert 

max_col5=vfx2_sh.cells(3, "F").end('right').column
vfx2_sh.api.columns(max_col5+1).insert 

max_col6=mal_sh.cells(3, "F").end('right').column
mal_sh.api.columns(max_col6+1).insert 

max_col7=vtl_sh.cells(3, "F").end('right').column
vtl_sh.api.columns(max_col7+1).insert 

max_col8=mal2_sh.cells(3, "F").end('right').column
mal2_sh.api.columns(max_col8+1).insert 

max_col9=vtl2_sh.cells(3, "F").end('right').column
vtl2_sh.api.columns(max_col9+1).insert 

max_col10=mt5_sh.cells(3, "F").end('right').column
mt5_sh.api.columns(max_col10+1).insert 



s_c_sh.range(f'D{s_c_sh.cells(4, "D").end("down").row+1}').value = s_c
m_c_sh.range(f'D{m_c_sh.cells(4, "D").end("down").row+1}').value = m_c
m_c_sh.range(f'N{m_c_sh.cells(4, "N").end("down").row+1}').value = m_c1
for i in range(len(vfx1)):
    vfx_sh.range(f"{chr(ord('@')+max_col4+1)}{i+3}").value = vfx1[i]

for i in range(len(vfx2)):
    vfx2_sh.range(f"{chr(ord('@')+max_col5+1)}{i+3}").value = vfx2[i]

for i in range(len(mal1)):
    mal_sh.range(f"{chr(ord('@')+max_col6+1)}{i+3}").value = mal1[i]
    
for i in range(len(vtl1)):
    vtl_sh.range(f"{chr(ord('@')+max_col7+1)}{i+3}").value = vtl1[i]

for i in range(len(mal2)):
    mal2_sh.range(f"{chr(ord('@')+max_col8+1)}{i+3}").value = mal2[i]

for i in range(len(vtl2)):
    vtl2_sh.range(f"{chr(ord('@')+max_col9+1)}{i+3}").value = vtl2[i]

for i in range(len(mt5)):
    mt5_sh.range(f"{chr(ord('@')+max_col10+1)}{i+3}").value = mt5[i]

res_sh.range(11,max_col1+1).value = res[0]
res_sh.range(12,max_col1+1).value = res[1]
res_sh.range(15,max_col1+1).value = res[2]
res_sh.range(16,max_col1+1).value = res[3]
res_sh.range(32,max_col1+1).value = res[4]
res_sh.range(33,max_col1+1).value = res[5]
res_sh.range(36,max_col1+1).value = res[6]
res_sh.range(37,max_col1+1).value = res[7]
res_sh.range(53,max_col1+1).value = res[8]
res_sh.range(54,max_col1+1).value = res[9]
res_sh.range(57,max_col1+1).value = res[10]
res_sh.range(58,max_col1+1).value = res[11]

res_sh.range(11,max_col2+1).value = res[12]
res_sh.range(12,max_col2+1).value = res[13]
res_sh.range(15,max_col2+1).value = res[14]
res_sh.range(16,max_col2+1).value = res[15]
res_sh.range(32,max_col2+1).value = res[16]
res_sh.range(33,max_col2+1).value = res[17]
res_sh.range(36,max_col2+1).value = res[18]
res_sh.range(37,max_col2+1).value = res[19]
res_sh.range(53,max_col2+1).value = res[20]
res_sh.range(54,max_col2+1).value = res[21]
res_sh.range(57,max_col2+1).value = res[22]
res_sh.range(58,max_col2+1).value = res[23]

res_sh.range(2,max_col1+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(23,max_col1+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(44,max_col1+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(2,max_col2+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(23,max_col2+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(44,max_col2+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(2,max_col3+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(23,max_col3+1).value = '{}/{}/{}'.format(TY,TM,TD)
res_sh.range(44,max_col3+1).value = '{}/{}/{}'.format(TY,TM,TD)
vfx_sh.range(2,max_col4+1).value = '{}/{}/{}'.format(TY,TM,TD)
vfx2_sh.range(2,max_col5+1).value = '{}/{}/{}'.format(TY,TM,TD)
mal_sh.range(2,max_col6+1).value = '{}/{}/{}'.format(TY,TM,TD)
vtl_sh.range(2,max_col7+1).value = '{}/{}/{}'.format(TY,TM,TD)
mal2_sh.range(2,max_col8+1).value = '{}/{}/{}'.format(TY,TM,TD)
vtl2_sh.range(2,max_col9+1).value = '{}/{}/{}'.format(TY,TM,TD)
mt5_sh.range(2,max_col10+1).value = '{}/{}/{}'.format(TY,TM,TD)

for i in [4,5,6,8,9,13,17,19,25,26,27,29,30,34,38,40,46,47,48,50,51,55,59]:
    for n in [max_col1,max_col2,max_col3]:
        res_sh.range(f"{colnum_string(n)}{i}").api.AutoFill(res_sh.range(f"{colnum_string(n)}{i}:{colnum_string(n+1)}{i}").api,AutoFillType.xlFillDefault)    

for i in [3,7,10,16,17,18,26,27,28,35,38,39,40,47,48,49,54,55,56,60,61,62,69,70,72,73]:
    vfx_sh.range(f"{colnum_string(max_col4)}{i}").api.AutoFill(vfx_sh.range(f"{colnum_string(max_col4)}{i}:{colnum_string(max_col4+1)}{i}").api,AutoFillType.xlFillDefault)    

for i in [3,7,10,16,17,18,26,27,28,35,38,39,40,47,48,49,54,55,56]:
    for s in [vfx2_sh,mal_sh,mal2_sh]:
        s.range(f"{colnum_string(max_col5)}{i}").api.AutoFill(s.range(f"{colnum_string(max_col5)}{i}:{colnum_string(max_col5+1)}{i}").api,AutoFillType.xlFillDefault)    

for i in [3,7,10,16,17,18,24,25,26,33,36,37,38,45,46,47,52,53,54]:
    for s in [vtl_sh,vtl2_sh]:
        s.range(f"{colnum_string(max_col7)}{i}").api.AutoFill(s.range(f"{colnum_string(max_col7)}{i}:{colnum_string(max_col7+1)}{i}").api,AutoFillType.xlFillDefault)    
    
for i in [3,7,9,15,16,17,25,26,27,32,35,36,37,44,45,46,51,52,53]:
    mt5_sh.range(f"{colnum_string(max_col10)}{i}").api.AutoFill(mt5_sh.range(f"{colnum_string(max_col10)}{i}:{colnum_string(max_col10+1)}{i}").api,AutoFillType.xlFillDefault)    

for i in [11,12,15,16,32,33,36,37,53,54,57,58]:
    res_sh.range(f"{colnum_string(max_col3)}{i}").api.AutoFill(res_sh.range(f"{colnum_string(max_col3)}{i}:{colnum_string(max_col3+1)}{i}").api,AutoFillType.xlFillDefault)    

wb.save(r'\\192.168.1.20\Rmc\Data Analysis\Dropbox\Daily PNL\{}-{}\{}{}\PNL\VFX-PNL-{}{}{}.xlsx'.format(TY,TM,TM,TD,TY,TM,TD))
wb.close()
app.quit()   


# In[ ]:




