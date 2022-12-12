import requests
import math
import os
import openpyxl 
import sys
import time
import concurrent.futures
from openpyxl.styles import Font
from datetime import datetime


def generate_filename():
    now = datetime.now()
    xx = 'Output_' + str(now)[:-7] + '.xlsx'
    xx = xx.replace(' ', '_').replace(':', '-')
    return xx

filename = 'Output.xlsx'
if os.path.exists(filename):
    filename = generate_filename()


wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Name', "Email", "Telephone", "Country"])
for cell in ws["1:1"]:
    cell.font = Font(bold=True)
page_size = 50



def scrape_ambiente(page):
    page_size = 50
    url = f'https://exhibitorsearch.messefrankfurt.com/service/esb/2.1/search/exhibitor?language=en-GB&q=&orderBy=name&pageNumber={page}&pageSize={page_size}&showJumpLabels=true&findEventVariable=AMBIENTE'
    res = requests.get(url).json()
    records = res['result']['hits']

    for record in records:
        name = record['exhibitor']['name']
        tel = record['exhibitor']['address']['tel']
        country = record['exhibitor']['address']['country']['label']
        email = record['exhibitor']['address']['email']
        row = [name, email, tel, country]
        ws.append(row)
    wb.save(filename)
    

scrape_ambiente(3)