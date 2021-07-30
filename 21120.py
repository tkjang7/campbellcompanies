import os, os.path
import win32com.client
import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd

#### time setsup ####
#####################
today_ = datetime.datetime.today().strftime("%m/%d/%Y")
dt_obj = datetime.datetime.today().strftime("%m/%d/%Y %H:%M:%S")
lastmonth = (datetime.datetime.today() + relativedelta(months=-1)).month
lastyear = (datetime.datetime.today() + relativedelta(months=-1)).year
twomonth = (datetime.datetime.today() + relativedelta(months=-2)).month
twoyear = (datetime.datetime.today() + relativedelta(months=-2)).year



#### win32com setup ####
########################

pathexist = os.path.exists('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab\\21120-{}.xlsm'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2)))

if pathexist:

    xl=win32com.client.DispatchEx('Excel.Application')
    xl.visible = True
    wb = xl.Workbooks.Open(Filename = 'S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab\\21120-{}.xlsm'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2))
    , ReadOnly=False)

    #### manipulation ####
    ######################
    ws = wb.Worksheets("Summary")
    ws.Cells(1,9).Value = 'Updated ' + dt_obj
    ws.Cells(2,8).Value = lastmonth
    ws.Cells(2,6).Value = lastyear

    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()


    ###### reconcilication by machine
    ######################################
    ######################################
    ws3 = wb.Worksheets('Estimates Data_')
    lastrow=ws3.UsedRange.Rows.Count

    li1=ws3.Range('A1:A{}'.format(lastrow)).Value


    ws4 = wb.Worksheets('TB as of now (machines)')
    lastrow=ws4.UsedRange.Rows.Count

    li2 = ws4.Range('D1:D{}'.format(lastrow)).Value



    machine1 = []

    for i in range(1,len(li1)):
        machine1.append(li1[i][0])


    machine2 = []

    for i in range(1,len(li2)):
        machine2.append(li2[i][0])



    ws2 = wb.Worksheets('discrepencies')
    ws2.AutoFilterMode = False
    ws2.Cells.Clear()

    idx = 2
    ws2.Cells(1,1).Value = 'Machine ID'
    ws2.Cells(1,2).Value = 'DBS'
    ws2.Cells(1,3).Value = 'CODA'
    ws2.Cells(1,4).Value = 'diff'
    machine_li=list(set(machine1+machine2)) # Unique machine IDs from both DBS and CODA
    for j in machine_li:

        ws2.Cells(idx,1).Value = j
        idx += 1


    ### Formula autofill
    ws2.Range('B2').Formula = '=SUMIF(Query4[Machine ID],discrepencies!A2,Query4[Estimate Amount])'

    ws2.Range('C2').Formula = '=SUMIF(Query16[REF2],discrepencies!A2,Query16[SumOfVALUEDOC])'

    ws2.Range('D2').Formula = '=B2+C2'

    lastrow = ws2.UsedRange.Rows.Count
    sourceRange = ws2.Range('B2:D2')
    fillRange = ws2.Range('B2:D{}'.format(lastrow))

    sourceRange.AutoFill(Destination = fillRange)



    ws2.UsedRange.NumberFormat = "#,##0.00_);(#,##0.00)"

    pdlist =[]
    dt = ws2.UsedRange.Value

    for i in range(1,len(dt)):
        i=dt[i]
        pdlist.append(list(i))

    pd_dt = pd.DataFrame(pdlist, columns=dt[0])

    str(pd_dt['Machine ID'][0])+'U' in list(pd_dt['Machine ID'])

    Olist = [] # original machine IDs
    Ulist = [] # machines IDs with U at the end


    for i in range(len(pdlist)):   ### If there are duplicate machine IDs (xxxxxxU and xxxxxx), highlight them.

        if str(pd_dt['Machine ID'][i])+'U' in list(pd_dt['Machine ID']):

            #Index of +'U'
            idx = list(pd_dt['Machine ID']).index(pd_dt['Machine ID'][i]+'U')


            # if the difference of diffrence is 0...
            if pd_dt['diff'][i] + pd_dt['diff'][idx] == 0:
                print(pd_dt['Machine ID'][i], i, idx)


                Olist.append(i)
                Ulist.append(idx)



    for o in Olist:
        oRow = o+2
        ws2.Range('A{}:D{}'.format(oRow,oRow)).Interior.Color = '&H85CA56'


    for u in Ulist:
        uRow = u+2
        ws2.Range('A{}:D{}'.format(uRow,uRow)).Interior.Color = '&H85CA56'

    ws2.UsedRange.AutoFilter(4, Criteria1='> 0.01', Operator=2, Criteria2='<-0.01')

    ##########################################################
    ##########################################################

    wb.Save()

    wb.Close(1)
    xl.Quit()

else:
    try:
        #### Makedirs ####
        os.makedirs('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2)), exist_ok=True)

        xl=win32com.client.DispatchEx('Excel.Application')
        xl.visible = True
        wb = xl.Workbooks.Open(Filename = 'S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab\\21120-{}.xlsm'.format(str(twoyear), str(twoyear), str(twomonth).zfill(2), str(twomonth).zfill(2))
        , ReadOnly=False)

        #### manipulation ####
        ######################
        ws = wb.Worksheets("Summary")
        ws.Cells(1,9).Value = 'Updated ' + dt_obj
        ws.Cells(2,8).Value = lastmonth
        ws.Cells(2,6).Value = lastyear

        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()


        wb.SaveAs('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab\\21120-{}.xlsm'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2), str(lastmonth).zfill(2)))

        wb.Close(0)
        xl.Quit()

    except:
        exit()
