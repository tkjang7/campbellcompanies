import os, os.path
import win32com.client
import datetime
from dateutil.relativedelta import relativedelta

#### time setsup ####
#####################
today_ = datetime.datetime.today().strftime("%m/%d/%Y")
dt_obj = datetime.datetime.today().strftime("%m/%d/%Y %H:%M:%S")
# lastmonth = (datetime.datetime.today() + relativedelta(months=0)).month
# lastyear = (datetime.datetime.today() + relativedelta(months=0)).year
# twomonth = (datetime.datetime.today() + relativedelta(months=-2)).month
# twoyear = (datetime.datetime.today() + relativedelta(months=-2)).year

vpath = "I:\\projects\\Machine Inventory - Tatyana\\prod\\13XXX-mm Machine Inventory.xlsm"


#### win32com setup ####
########################

pathexist = os.path.exists(vpath)

if pathexist:

    xl=win32com.client.Dispatch('Excel.Application')
    xl.visible = True
    wb = xl.Workbooks.Open(Filename = vpath
    , ReadOnly=False)

    #### manipulation ####
    ######################
    ws = wb.Worksheets("Recon_")
    ws.Cells(28,1).Value = 'Updated ' + dt_obj


    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()


    #### manipulation end ####
    ##########################
    wb.Save()


    wb.Close(1)
    xl.Quit()






    # ######### Send Email to Tatyana #########
    # #########################################
    # outlook = win32com.client.Dispatch('Outlook.Application').CreateItem(0)#GetNamespace("MAPI")
    # # accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    #
    # outlook.To = 'tatyanak@campbellcompanies.com'
    # outlook.Subject = 'Automated Email: 13xxx Machine Balancing'
    # outlook.HTMLBody = 'Please find the attachment updated as of {}'. format(dt_obj)
    #
    #
    # outlook.Attachments.Add(vpath)
    # # outlook.Display(True)
    #
    # outlook.Send()

else:
    print(pathexist)


try:
    #### Makedirs ####
    # os.makedirs('S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\{}\{}-{}\\21100s - Other Current Liab'.format(str(lastyear), str(lastyear), str(lastmonth).zfill(2)), exist_ok=True)

    xl=win32com.client.DispatchEx('Excel.Application')
    xl.visible = True
    wb = xl.Workbooks.Open(Filename = "S:\\01 Wheeler Files\\Departments\\General Accounting\\01 Month End Recons\\13XXX-mm Machine Inventory.xlsm"
    , ReadOnly=False)

    #### manipulation ####
    ######################
    ws = wb.Worksheets("Recon_")
    ws.Cells(28,1).Value = 'Updated ' + dt_obj


    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()


    wb.Save()

    wb.Close(1)
    xl.Quit()

except:
    exit()
