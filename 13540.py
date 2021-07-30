import os, os.path
import win32com.client
import datetime
from dateutil.relativedelta import relativedelta

#### time setsup ####
#####################
today_ = datetime.datetime.today().strftime("%m/%d/%Y")
dt_obj = datetime.datetime.today().strftime("%m/%d/%Y %H:%M:%S")
currentmonth = datetime.datetime.today().month
lastmonth = (datetime.datetime.today() + relativedelta(months=-1)).month
lastyear = (datetime.datetime.today() + relativedelta(months=-1)).year
prevmonth = (datetime.datetime.today() + relativedelta(months=-4)).month     #### <---- IMPORTANT: Schedule quarterly update
prevyear = (datetime.datetime.today() + relativedelta(months=-4)).year





#### win32com setup ####
########################

pathexist = os.path.exists('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\13000s - Inventory\\13540-{}.xlsx'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2)))


if pathexist:

    xl=win32com.client.Dispatch('Excel.Application')
    xl.visible = True
    wb = xl.Workbooks.Open(Filename = 'S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\13000s - Inventory\\13540-{}.xlsx'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2))
    , ReadOnly=False)

    #### manipulation ####
    ######################
    ws = wb.Worksheets("Summary")
    ws.Cells(5,10).Value = 'Updated ' + dt_obj
    ws.Cells(2,8).Value = lastmonth
    ws.Cells(2,6).Value = lastyear

    ws2 = wb.Worksheets("Invoiced WO")
    ws2.Cells(3,2).Value = currentmonth



    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()


    wb.Save()

    wb.Close(1)
    xl.Quit()

else:
    try:
        #### Makedirs ####
        os.makedirs('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\13000s - Inventory'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2)), exist_ok=True)

        xl=win32com.client.Dispatch('Excel.Application')
        xl.visible = True
        wb = xl.Workbooks.Open(Filename = 'S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\13000s - Inventory\\13540-{}.xlsx'.format(str(prevyear), str(prevyear), str(prevmonth).zfill(2), str(prevmonth).zfill(2))
        , ReadOnly=False)

        #### manipulation ####
        ######################
        ws = wb.Worksheets("Summary")
        ws.Cells(1,9).Value = 'Updated ' + dt_obj
        ws.Cells(2,8).Value = lastmonth
        ws.Cells(2,6).Value = lastyear


        ws2 = wb.Worksheets("Invoiced WO")
        ws2.Cells(3,2).Value = currentmonth

        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()


        wb.SaveAs('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\13000s - Inventory\\13540-{}.xlsx'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2)))

        wb.Close(0)
        xl.Quit()

    except:
        exit()
