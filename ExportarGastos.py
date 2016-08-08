#!/usr/bin/env python

import calendar
import datetime
import sys
import logging
import gspread
import glob
import os
from oauth2client.service_account import ServiceAccountCredentials
from BeautifulSoup import BeautifulSoup
from itertools import groupby

__author__ = 'rdelgado'

''' Parse and get the credit transactions data from the bank registry '''
def getCreditTransactionList(path):
    bankData = open(path, 'r').read()
    soup = BeautifulSoup(bankData)
    table = soup.find("table", {"id" : "detalle_transacciones"})
    readedRows = []
    for row in table.findAll('tr')[2:]:
        cols = row.findAll('td')
        try:
            if float(cols[2].div.span.string.replace(",", "")) < 0:
                continue;
            readedRows.append([parseBankDate(cols[0].div.span.string.strip()), cols[1].div.span.string.strip(), 'Credito', float(cols[2].div.span.string.replace(",", ""))])
        except Exception as e:
            print(e)
    return readedRows

''' Choose the category for the credit payment from the debit transaction data. '''
def setPaymentCategory(concept):
    if 'PAGO TARJETA' in concept:
        return 'Pago credito'
    else:
        return 'Debito'


''' Parse and get the debit transaction data from the bank registry '''
def getDebitTransactionList(path):
    bankData = open(path, 'r').read()
    soup = BeautifulSoup(bankData)
    table = soup.find("table", {"class" : "transaction"})
    readedRows = []
    for row in table.findAll('tr'):
        cols = row.findAll('td')
        try:
            if float(cols[2].div.span.string.replace(",", "")) < 0:
                continue;
            readedRows.append([parseBankDate(cols[0].div.span.string.strip()), cols[1].div.span.string.strip(), setPaymentCategory(cols[1].div.span.string.strip()), float(cols[2].div.span.string.replace(",", ""))])
        except Exception as e:
            print(e)
    return readedRows

def key(item):
    return unicode(item[2])

''' Group by category '''
def groupRecors(records):
    category = lambda x: unicode(x[2])
    sortedRecords = sorted(records, key=category)
    grouped = groupby(sortedRecords, category)
    records = []
    for key, values in grouped:
        sum = 0
        for v in values:
            sum = sum + float(v[3]) * -1;
        records.append([key, sum])
    return records


''' Parse date from bank data'''
def parseBankDate(date):
    year = date[6:]
    month = date[3:-5]
    day = date[:2]
    return month + "/" + day + "/" + year

''' Parse date '''
def parseDate(rawDate):
    year = rawDate[:4]
    month = rawDate[4:-2]
    day = rawDate[6:]
    return month + "/" + day + "/" + year

''' Upload the data to Google Docs '''
def extractSpendDataFromGoogleDocs():
    # Open a worksheet from spreadsheet with one shot
    wks = gc.open("AndroMoney").sheet1
    records = wks.get_all_values()[2:]
    cleanRecords = []
    for row in records:
        if len(row[7]) == 0:
            cleanRecords.append([parseDate(row[5]), row[8], row[3], float(row[2])])
    return cleanRecords

''' Check if the tool must run '''
def mustRun(localPath):
    return len(glob.glob(localPath + '/*Debit.html')) > 0 or len(glob.glob(localPath + '/*Credit.html')) > 0


''' Main code of the tool '''
## Configuring logging facilities
reload(sys)
sys.setdefaultencoding('utf-8')
localPath = sys.argv[1]
print localPath
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%y %H:%M',
                    filename=localPath+'/ControlGastos.log',
                    filemode='w')
logger = logging.getLogger(__name__)


if not mustRun(localPath):
    logger.info('Theres no need to execute, Bye!')
    exit(1)

lastMonth = calendar.month_name[datetime.datetime.now().month - 1] # Minus one bacause we want the last month, not the current
currentYear = datetime.datetime.now().year

# Accesign to Google Spreadsheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(localPath+'/ControlGastos-e52eb4634fbe.json', scope)
gc = gspread.authorize(credentials)

# Open a worksheet from spreadsheet with one shot
wks = gc.open("Presupuesto anual "+str(currentYear)).worksheet(lastMonth)

# Reading the spend data from google
spendRecords = extractSpendDataFromGoogleDocs()

# Reading the spend data from Debit movements
for f in glob.glob(localPath + '/*Debit.html'):
    spendRecords.extend(getDebitTransactionList(f))
    os.remove(f)

# Reading the spend data from Credit movements
for f in glob.glob(localPath + '/*Credit.html'):
    transactions = getCreditTransactionList(f)
    os.remove(f)

for index, row in enumerate(groupRecors(spendRecords)):
    wks.update_cell(4 + index, 4, str(row[0]))
    wks.update_cell(4 + index, 5, str(row[1]))

for index, row in enumerate(spendRecords):
    wks.update_cell(30 + index, 2, str(row[0]))
    wks.update_cell(30 + index, 3, unicode(str(row[2])))
    wks.update_cell(30 + index, 4, unicode(str(row[1])))
    wks.update_cell(30 + index, 5, str(row[3]))

if len(transactions) > 0:
    for index, row in enumerate(transactions):
        wks.update_cell(30 + index, 7, str(row[0]))
        wks.update_cell(30 + index, 8, unicode(str(row[1])))
        wks.update_cell(30 + index, 9, unicode(str(row[3])))








