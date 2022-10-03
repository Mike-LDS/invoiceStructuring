#!/usr/bin/env python3

import PySimpleGUI as sg
import os.path
import csv
import datetime

client_2_students = []
invoice_2_client = []
invoice_detialed = []
invoice_sum = []
error= []
adhoc = []

header_d = ['Invoice Number', 'Customer', 'Invoice Date', 'Due Date', 'Item', 'Quantity', 'Rate', 'Total', 'Item Tax Code', 'Exc/Inc of Tax', 'Service Date']
header_s = ['Invoice Number', 'Customer', 'Invoice Date', 'Due Date', 'Item', 'Item Description','Quantity', 'Rate', 'Total', 'Item Tax Code']

sg.theme("Dark2")
layout = [
    [sg.T("")],
    [sg.Text("Choose Student Export: "), sg.Input(), sg.FileBrowse(file_types=(('CSV Files', '*.CSV'),), key='-USERS-')],
    [sg.Text('Go to TC -> System -> Exports -> Click on Users -> Select Students from drop down menu -> Click Submit')],
    [sg.T("")],
    [sg.Text("Choose Invoice Export:  "), sg.Input(), sg.FileBrowse(file_types=(('CSV Files', '*.CSV'),), key='-INVOICES-')],
    [sg.Text('Go to TC -> System -> Exports -> Click on Accounting-> Select date range & Invoices (Detialed) -> Click Submit')],
    [sg.T("")],
    [sg.Text("Starting Invoice Number:"), sg.InputText(size=(10,6), key='-NUM-')],
    [sg.T("")],
    [sg.Button("RUN"),sg.Text("", key='-TEXT-')]]

# Building Window
window = sg.Window('Invoice Structuring', layout, size=(600,270))
    
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="EXIT":
        break
        
    elif event == "RUN":       
# Reading TC Exports
        try:
            inv_num = int(values['-NUM-'])
        # Defining Link Between Parent & Student
            with open(values['-USERS-'], newline='', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    i = 1
                    line = [(row['Last name'] + ', ' + row['First name']), row['First name'], row['Client ID']]
                    client_2_students.append(line)

        # Defining Link Between Parent & Invoice
            with open(values['-INVOICES-'], newline='', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if row['InvoiceNumber'] in (client[1] for client in invoice_2_client):
                        continue
                    else:
                        line = [row['ID'], row['InvoiceNumber'], row['InvoiceDate'], row['DueDate']]
                        invoice_2_client.append(line)

        # Creating Line By Line Array
            with open(values['-INVOICES-'], newline='', encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if 'explicit' in row['Description'].lower():
                        item = 'Explicit Instruction'
                    elif 'support' in row['Description'].lower():
                        item = 'Homework Support'
                    elif 'rise at school' in row['Description'].lower():
                        item = 'RISE AT School'
                    elif 'slp' in row['Description'].lower():
                        item = 'SLP'
                    elif 'rise now' in row['Description'].lower():
                        item = 'RISE Now'
                    elif 'rise team' in row['Description'].lower():
                        item = 'RISE Team'
                    elif 'ktea-3' in row['Description'].lower():
                        item = 'KTEA-3 Assessment'
                    elif 'social language group' in row['Description'].lower():
                        item = 'Social Language Group'
                    elif 'summer tutoring' in row['Description'].lower():
                        item = 'Summer Tutoring'
                    elif 'summer rise' in row['Description'].lower():
                        item = 'Summer RISE'
                    elif 'intensive' in row['Description'].lower():
                        item = 'Summer RISE'
                    elif 'summer camp' in row['Description'].lower():
                        item = 'Summer Camp'
                    elif 'spring break camp' in row['Description'].lower():
                        item = 'Spring Break Camp'
                    elif 'early risers' in row['Description'].lower():
                        item = 'Early RISErs'
                    elif 'intake' in row['Description'].lower():
                        item = 'Intake and Assessment Fee'
                    elif 'onboarding' in row['Description'].lower():
                        item = 'Intake and Assessment Fee'
                    elif 'social language' in row['Description'].lower():
                        item = 'SLP'
                    elif 'family coaching' in row['Description'].lower():
                        item = 'Family Coaching'
                    elif 'p.e.e.r.s' in row['Description'].lower():
                        item = 'P.E.E.R.S'
                    elif 'homework club' in row['Description'].lower():
                        item = 'Homework Club'
                    elif 'neuropsychological' in row['Description'].lower():
                        item = 'Neuropsychological Assessment Fees'
                    elif 'parent workshop' in row['Description'].lower():
                        item = 'Parent Workshop'
                    elif 'wamms' in row['Description'].lower():
                        item = 'WAMMS'
                    else:
                        item = 'ERROR'
                    if row['TrackingOption1'] == "lesson":
                        cost = float(row['Quantity'])*float(row['UnitAmount'])
                        for client in invoice_2_client:
                            if row['InvoiceNumber'] in client[1]:
                                client_ID = client[0]
                                match = False
                                for student in client_2_students:
                                    if client_ID == student[2]:
                                        for i in range(len(student[1].split(' ',-1))):
                                            if student[1].split(' ',-1)[i] in row['Description']:
                                                customer = student[0]
                                                line = [row['InvoiceNumber'], customer, client[2], client[3], item, row['Quantity'], row['UnitAmount'], cost,'E','Exclusive',row['StartDate']]
                                                invoice_detialed.append(line)
                                                match = True
                                if match == False:
                                    for student in client_2_students:
                                        if client_ID == student[2]:
                                            customer = student[0]
                                            line = [row['InvoiceNumber'], customer, client[2], client[3], item, row['Quantity'], row['UnitAmount'], cost,'E','Exclusive',row['StartDate']]
                                            invoice_detialed.append(line)
                                            match = True
                                            break
                                if match == False:
                                    line = [row['InvoiceNumber'], client_ID, client[2], client[3], item, row['Quantity'], row['UnitAmount'], cost,'E','Exclusive',row['StartDate']]
                                    invoice_detialed.append(line)

                    elif row['TrackingOption1'] == "adhoc_charge":
                        for client in invoice_2_client:
                            if row['InvoiceNumber'] in client[1]:
                                client_ID = client[0]
                                match = False
                                for student in client_2_students:
                                    if client_ID == student[2]:
                                        for i in range(len(student[1].split(' ',-1))):
                                            if student[1].split(' ',-1)[i] in row['Description']:
                                                customer = student[0]
                                                line = ['', customer, row['InvoiceDate'], row['DueDate'], item, row['Description'], row['Quantity'], row['UnitAmount'], row['UnitAmount'],'E','Exclusive']
                                                adhoc.append(line)
                                                match = True
                                if match == False:
                                    for student in client_2_students:
                                        if client_ID == student[2]:
                                            customer = student[0]
                                            line = ['', customer, row['InvoiceDate'], row['DueDate'], item, row['Description'], row['Quantity'], row['UnitAmount'], row['UnitAmount'],'E','Exclusive']
                                            adhoc.append(line)
                                            match = True
                                            break
                                if match == False:
                                    line = ['', client_ID, row['InvoiceDate'], row['DueDate'], item, row['Description'], row['Quantity'], row['UnitAmount'], row['UnitAmount'],'E','Exclusive']
                                    adhoc.append(line)
                                    
        # Grouping Invoices
            for row in invoice_detialed:
                match = False
                for line in invoice_sum:
                    if line[header_s.index('Customer')] == row[header_d.index('Customer')] and line[header_s.index('Item')] == row[header_d.index('Item')] and float(line[header_s.index('Rate')]) == float(row[header_d.index('Rate')]):
                        line[header_s.index('Quantity')] = float(line[header_s.index('Quantity')]) + float(row[header_d.index('Quantity')])
                        line[header_s.index('Total')] = float(line[header_s.index('Quantity')])*float(line[header_s.index('Rate')])
                        line[header_s.index('Item Description')] = line[header_s.index('Item Description')]+ ', ' + row[header_d.index('Service Date')]
                        match = True
                if match == False:
                    line = ['', row[header_d.index('Customer')], row[header_d.index('Invoice Date')], row[header_d.index('Due Date')], row[header_d.index('Item')],
                            row[header_d.index('Item')]+' on '+row[header_d.index('Service Date')], float(row[header_d.index('Quantity')]), float(row[header_d.index('Rate')]),
                            float(row[header_d.index('Quantity')])*float(row[header_d.index('Rate')]), row[header_d.index('Item Tax Code')]]
                    invoice_sum.append(line)
                    
        # Adjust GL Codes 
            for row in invoice_sum:
                if row[header_s.index('Item')] == 'Social Language Group':
                    row[header_s.index('Item')] = 'SLP'
                elif row[header_s.index('Item')] == 'Summer Tutoring':
                    row[header_s.index('Item')] = 'Homework Support'
            for row in adhoc:
                if row[header_s.index('Item')] == 'Summer Camp' or row[header_s.index('Item')] == 'Spring Break Camp':
                    row[header_s.index('Item')] = 'CAMPS'

        # Adding Invoice Number:
            for row in invoice_sum:
                match = False
                for line in invoice_sum:
                    if line[header_s.index('Customer')] == row[header_s.index('Customer')] and line[header_s.index('Invoice Number')] != '':
                        row[header_s.index('Invoice Number')] = line[header_s.index('Invoice Number')]
                        match = True
                if match == False:
                    row[header_s.index('Invoice Number')] = inv_num
                    inv_num = inv_num + 1

            for row in adhoc:
                match = False
                for line in adhoc:
                    if line[header_s.index('Customer')] == row[header_s.index('Customer')] and line[header_s.index('Invoice Number')] != '':
                        row[header_s.index('Invoice Number')] = line[header_s.index('Invoice Number')]
                        match = True
                if match == False:
                    row[header_s.index('Invoice Number')] = inv_num
                    inv_num = inv_num + 1

        # Clearing Previous Invoices
            path = os.getcwd().split('/dist',1)[0]
            x = datetime.datetime.now()
            file = '/Invoices-' + x.strftime("%b") + x.strftime("%d") + x.strftime("%Y")+ '.csv'
            f = open(path+file, 'w', newline='', encoding="utf-8")
            f.close()
            
        # Writing Invoices to CSV
            with open(path+file, 'w', newline='', encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(header_s)
                for row in invoice_sum:
                    writer.writerow(row)

        # Clearing Previous Invoices
            x = datetime.datetime.now()
            file = '/AdhocCharges-' + x.strftime("%b") + x.strftime("%d") + x.strftime("%Y")+ '.csv'
            f = open(path+file, 'w', newline='', encoding="utf-8")
            f.close()
            
        # Writing Invoices to CSV
            with open(path+file, 'w', newline='', encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(header_s)
                for row in adhoc:
                    writer.writerow(row)

            window['-TEXT-'].update("INVOICES ARE READY")

        except Exception as e:
            sg.Popup(e, title='ERROR')

window.close()
