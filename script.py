import requests
import math
import os
import openpyxl 
import sys
import time
from openpyxl.styles import Font
from datetime import datetime


start = time.time()
total_records = 0
records_scraped = 0
print('Starting The Script...')
time.sleep(2)
print("Finding the total records...")

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
page_size = 200


def get_total_pages():
    global total_records
    url = 'https://exhibitorsearch.messefrankfurt.com/service/esb/2.1/search/exhibitor?language=en-GB&q=&orderBy=name&pageNumber=1&pageSize=1&showJumpLabels=true&findEventVariable=AMBIENTE'
    res = requests.get(url).json()
    records = res['result']['metaData']['hitsTotal']
    total_records = records
    return math.ceil(records/200)
    

def scrape_ambiente(page):
    global records_scraped, page_size
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
        records_scraped +=1 
    wb.save(filename)
    

total_pages = get_total_pages()
print(f'Found {total_records} Records..')
time.sleep(2)
print('Scraping the records...')

i = 1
while i <= total_pages:
    scrape_ambiente(i)
    print(f'Scraped {records_scraped} records...', end='\r')
    i += 1


end = time.time()
print(f'Completed in {int(end - start)} seconds.')
print(f'Please see the file {filename}')
time.sleep(4)
exit()