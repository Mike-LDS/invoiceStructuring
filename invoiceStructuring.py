import requests
import json
import pandas as pd
import time
import datetime
import PySimpleGUI as sg
import os.path


#------------------------------------------------ Building Window --------------------------------------------------------#
sg.theme("Dark2")
layout = [
    [sg.T("")],
    [sg.Text("Date Range:"), sg.InputText(size=(20,20), key='-START-'),sg.Text(" to ") ,sg.InputText(size=(20,20), key='-END-')],
    [sg.Text("Date Format should be YYYY-MM-DD (e.g. 2023-05-30)")],
    [sg.T("")],
    [sg.Text("Starting Invoice Number:"), sg.InputText(size=(20,20), key='-NUM-')],
    [sg.T("")],
    [sg.Button("RUN"),sg.Text("", key='-TEXT-')]]

window = sg.Window('Invoice Structuring', layout, size=(500,200))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="EXIT":
        break
#------------------------------------------------ Processing Invoices --------------------------------------------------------#
    elif event == "RUN":
        try:
            lessons_export = pd.DataFrame({'Invoice Number': [], 'Customer':[],'Invoice Date':[],'Due Date':[],'Item':[], 'Item Description':[], 'Quantity':[], 'Rate':[], 'Total':[], 'Item Tax Code':[]})
            charges_export = pd.DataFrame({'Invoice Number': [], 'Customer':[],'Invoice Date':[],'Due Date':[],'Item':[], 'Item Description':[], 'Quantity':[], 'Rate':[], 'Total':[], 'Item Tax Code':[]})
            service_list = {}

            headers = {
              'Authorization': 'Token token=st_live_fvXWm27jS_Q_BRtwYWH5lg',
              'Content-Type': 'application/json'
            }

            inv_num = int(values['-NUM-'])

            #----------------- Creating a Service List ---------------------------------------#
            s_url = "https://api.teachworks.com/v1/services?per_page=50"


            response = requests.request("GET", s_url, headers=headers, data={})
            service_info = response.json()
            for i in range(0,len(service_info)):
                service_list[service_info[i]["id"]] = service_info[i]["name"]

            #----------------- Pulling Invoices ---------------------------------------#

            url = "https://api.teachworks.com/v1/invoices?per_page=80&page="

            payload= json.dumps({
             "date":{"gte":values['-START-'], "lte":values['-END-']}
            })

            j = 1
            num = 0
            flag = False
            while flag == False:
                window['-TEXT-'].update("Pulling Invoices from Page "+str(j))
                response = requests.request("GET", url+str(j), headers=headers, data=payload)
                invoices = response.json()
                if invoices == ['Rate Limit Exceeded']:
                    time.sleep(0.1)
                    response = requests.request("GET", url+str(j), headers=headers, data=payload)
                    invoices = response.json()
                for i in range(0,len(invoices)):
                    fam_id = invoices[i]["customer_id"]
                    date = invoices[i]["date"]
                    due = invoices[i]["due_date"]
                    
            #----------------- Compiling Individual Lessons ---------------------------------------#
                    for lesson in invoices[i]["lessons"]:
                        cusomter = lesson['student_id']
                        item = service_list[lesson['service_id']]
                        item_d = item + ' on ' +(lesson['description'].split('\r\n',1)[1]).split(' ',1)[0]
                        rate = lesson['invoice_unit_price']
                        total = lesson['invoice_amount']
                        duration = (lesson['description'].split('\r\n',1)[1]).split(' ',1)[1]
                        start_time = datetime.datetime.strptime(duration.split(' - ',-1)[0], '%I:%M %p')
                        end_time = datetime.datetime.strptime(duration.split(' - ',-1)[1], '%I:%M %p')
                        quantity = abs((end_time - start_time).total_seconds())/3600

                        line = pd.DataFrame({'Customer':str(cusomter),'Invoice Date':date,'Due Date':due,'Item':str(item), 'Item Description':item_d, 'Quantity':quantity, 'Rate':rate, 'Total':total,
                                             'Item Tax Code':'E'}, index=[num])
                        lessons_export = pd.concat([lessons_export,line])
                        num = num + 1

            #----------------- Compiling Individual Charges ---------------------------------------#
                    for charge in invoices[i]["charges"]:
                        quantity = charge["quantity"]
                        rate = charge["unit_price"]
                        total = charge["amount"]
                        item_d = charge["description"]
                        item = charge["title"]
                        
                        inv_num = inv_num + 1
                        line = pd.DataFrame({'Invoice Number': 'INV-'+ str(inv_num), 'Customer':fam_id,'Invoice Date':date,'Due Date':due,'Item':item, 'Item Description':item_d, 'Quantity':quantity, 'Rate':rate, 'Total':total,
                                             'Item Tax Code':'E'}, index=[num])
                        charges_export = pd.concat([charges_export,line])
                        num = num + 1

                j = j + 1
                if invoices == [] or j == 100000:
                    flag = True

            #--------------------------- Adding INV # for Lessons ----------------------------------#
            students = lessons_export['Customer'].unique()
            families = charges_export['Customer'].unique()

            for student in students:
                inv_num = inv_num + 1
                lessons_export.loc[(lessons_export['Customer'] == student, 'Invoice Number')] = 'INV-'+str(inv_num)
                
            #----------------- Replacing the Customer Name (Lessons)------------------------------#

            window['-TEXT-'].update("Formattting Lesson Invoices")
            for student in students:
                stu_url = "https://api.teachworks.com/v1/students/"+student
                response = requests.request("GET", stu_url, headers=headers, data={})
                stu_info = response.json()

                if stu_info == ['Rate Limit Exceeded']:
                    time.sleep(0.1)
                    response = requests.request("GET", stu_url, headers=headers, data={})
                    stu_info = response.json()

                if stu_info == ['Not Found']:
                    time.sleep(0.1)
                    ind_url = "https://api.teachworks.com/v1/customers/"+student
                    response = requests.request("GET", ind_url, headers=headers, data={})
                    stu_info = response.json()
                    
                customer = stu_info['last_name'] +', '+stu_info['first_name']
                lessons_export = lessons_export.replace({'Customer':student}, {'Customer':customer}, regex=False)


            #----------------- Replacing the Customer Name (Charges)------------------------------#

            window['-TEXT-'].update("Formattting Charges")
            for family in families:
                stu_url = "https://api.teachworks.com/v1/students"
                payload = json.dumps({"customer_id":family})
                
                response = requests.request("GET", stu_url, headers=headers, data=payload)
                stu_info = response.json()
                
                if stu_info == ['Rate Limit Exceeded']:
                    time.sleep(0.1)
                    response = requests.request("GET", stu_url, headers=headers, data=payload)
                    stu_info = response.json()

                if stu_info == ['Not Found']:
                    time.sleep(0.1)
                    ind_url = "https://api.teachworks.com/v1/customers/"+student
                    response = requests.request("GET", ind_url, headers=headers, data={})
                    stu_info = response.json()

                if len(stu_info) == 1:
                    customer = stu_info[0]['last_name'] +', '+stu_info[0]['first_name']
                    charges_export = charges_export.replace({'Customer':family}, {'Customer':customer}, regex=False)
                    
            #----------------- Sorting and Exporting Invoices -----------------------------------#

            lessons_export = lessons_export.sort_values(by=['Invoice Number'])
            charges_export = charges_export.sort_values(by=['Invoice Number'])
            
            x = datetime.datetime.now()
            file_name = 'Lessons-' + x.strftime("%b") + x.strftime("%d") + x.strftime("%Y")+ '.csv'
            lessons_export.to_csv(file_name, index=False)
            file_name = 'Charges-' + x.strftime("%b") + x.strftime("%d") + x.strftime("%Y")+ '.csv'
            charges_export.to_csv(file_name,index=False)
            window['-TEXT-'].update("Invoices are Ready")

        except Exception as e:
            sg.Popup(e, title='ERROR')

window.close()
