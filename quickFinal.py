from __future__ import print_function

from googleapiclient import discovery
from google.oauth2 import service_account

from datetime import datetime,timedelta
import time
import random
import sys

from win32com.client import Dispatch
import pathlib

import json

def init_config():
    with open('./SHLabelPackage/config.json') as f:
        config = json.load(f)
    ID = config["ID"]
    return ID

def init_sheets(SPREADSHEET_ID):
    filePath = pathlib.Path('./SHLabelPackage/print.label')

    CLIENT_SECRET_FILE = './SHLabelPackage/keys.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            CLIENT_SECRET_FILE, scopes=SCOPES)

    service = discovery.build('sheets', 'v4', credentials=creds)
    spreadsheets = service.spreadsheets()

    resultDate = datetime.now()
    today_date = resultDate.strftime("%d/%m")
    add_sheets(SPREADSHEET_ID, today_date, spreadsheets)
    print("Added sheet")
    return service, spreadsheets, filePath

def add_sheets(gsheet_id, sheet_name, spreadsheets): #ONLY MAke new sheet if there isn't one
    try:
        print("win")
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                    }
                }
            }]
        }
        print("A")
        response = spreadsheets.batchUpdate(
            spreadsheetId=gsheet_id,
            body=request_body
        ).execute()
        print("A")
        return response
    except Exception as e:
        print("fail")
        print(e)

def init_printers():
    printerCOM = Dispatch('Dymo.DymoAddIn')
    printers = printerCOM.getDymoPrinters()
    theList = []
    temp = ''
    onlinePrinters = []
    for x in range(len(printers)):
        if printers[x] != '|':
            temp = temp + printers[x]
        else:
            theList.append(temp)
            temp = ''
    theList.append(temp)
    for x in theList:
        if printerCOM.isPrinterOnline(x) == True:
            onlinePrinters.append(x)

    if len(onlinePrinters) > 0:
        printerNumber = random.randrange(len(onlinePrinters))
    else:
        printerNumber = 0
    
    return printerCOM, onlinePrinters, printerNumber

def print_name(printerLabel, printerCOM, firstName, lastName):
    print("printing ", firstName, lastName)
    printerLabel.SetField('TEXT', firstName)
    printerLabel.SetField('TEXT_1', lastName)
    printerCOM.StartPrintJob()
    printerCOM.Print(1,False)
    printerCOM.EndPrintJob()

def call_sheets(service, SPREADSHEET_ID, cell):
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=cell).execute()
    print(result)
    if "values" in result.keys():
        # Read the dict
        firstName = result['values'][0][0]
        lastName = result['values'][0][1]
        return firstName, lastName
    else:
        return "BLANK!", ""

def main(service, spreadsheets, filePath, SPREADSHEET_ID):
    print("starting main")
    resultDate = datetime.now()
    today_date = resultDate.strftime("%d/%m")

    print("RUN first time code where it checks the first 1000 lines")
    cell = today_date + "!A1:C1000"
    print(cell)
    sheet = service.spreadsheets()
    print(SPREADSHEET_ID)
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=cell).execute()
    v = list(result.values())
    if len(v) > 2:
        currentRow = len(v[2]) + 1
    else:
        currentRow = 1
    print("Starting at row " + str(currentRow))

    while True:
        printerCOM, onlinePrinters, printerNumber = init_printers()

        myPrinter = onlinePrinters[printerNumber] #'DYMO LabelWriter 450 TURBO'
        print(myPrinter)
        printerCOM.selectPrinter(myPrinter)#myprinter
        printerCOM.Open2(filePath)
        printerLabel = Dispatch('Dymo.DymoLabels')
        # Set what row we're looking at
        cell = today_date + "!A" + str(currentRow) + ":C" + str(currentRow)
        print(cell)
        firstName, lastName = call_sheets(service, SPREADSHEET_ID, cell)
        if firstName != "BLANK!":
            if firstName == "Lamfam":
                print_name(printerLabel, printerCOM, "Harry", "Lam")
                print_name(printerLabel, printerCOM, "Shirley", "Lam")
                print_name(printerLabel, printerCOM, "Mikey", "Lam")
                print_name(printerLabel, printerCOM, "Elijah", "Lam")
                print_name(printerLabel, printerCOM, "Noah", "Lam")
                print_name(printerLabel, printerCOM, "Ezra", "Lam")
                currentRow = currentRow + 1
            elif firstName == "T":
                print_name(printerLabel, printerCOM, "Terence", "テレンス")
                currentRow = currentRow + 1
            elif firstName != "":
                print_name(printerLabel, printerCOM, firstName, lastName)
                currentRow = currentRow + 1
            else:
                currentRow = currentRow + 1
        else:
            print("NO DATA")

if __name__ == '__main__':
    ID = init_config()
    service, spreadsheets, filePath = init_sheets(ID)
    print("Initialized")
    main(service, spreadsheets, filePath, ID)