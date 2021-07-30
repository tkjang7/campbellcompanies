import pandas as pd
import win32com.client
import os
import datetime
from dateutil.relativedelta import relativedelta


y_ = (datetime.datetime.now() + relativedelta(months = -1)).year
m_ = str((datetime.datetime.now() + relativedelta(months = -1)).month).zfill(2)

path_ = "S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{y}\\{y}-{m}\\13000s - Inventory\\13XXX-{m} Parts Balancing.xlsx".format(y=y_, m=m_)
path_master = 'Z:\\Tableau WIP\\Excel Files\\WIP by Month - Monthly Parts Inventory_.csv'

### get last month's data
xl = win32com.client.Dispatch('Excel.Application')
xl.visible = True
wb = xl.Workbooks.Open(Filename = path_, ReadOnly = True, UpdateLinks=False)  #Open last month's reconciled Parts Balancing workbook

ws = wb.Worksheets("Inventory Summary")
lastrow=ws.UsedRange.Rows.Count

data_=ws.Range("B12:S{}".format(lastrow)).Value                               #Copy the data
df = pd.DataFrame(data_)
df = df.rename(columns=df.iloc[0]).drop(df.index[0])
df = df[pd.notna(df['Comp'])].reset_index(drop=True)

df = df.loc[:, df.columns.notnull()].reset_index(drop=True)

dat = []
for i in range(len(df)):
    dat.append('{month}-01-{year}'.format(month=m_, year=y_))

dat_ = pd.DataFrame(dat, columns = ["MonthYear"])
df=pd.concat([df, dat_], axis = 1)

wb.Close(SaveChanges=False)



### open master csv file and append and save
master = pd.read_csv(path_master)
comb = pd.concat([master, df], axis = 0)
print(comb)

comb.to_csv(path_master, index=0)

#
# wb_master = xl.Workbooks.Open(Filename = path_master, ReadOnly = True, UpdateLinks=False)
# ws_master = wb.Worksheets("WIP by Month - Monthly Parts In")
# lastrow_master=ws.UsedRange.Rows.Count
