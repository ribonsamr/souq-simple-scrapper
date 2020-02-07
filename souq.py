import re
import csv

import lxml
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

webpages = [{
    'link': 'https://deals.souq.com/eg-en/mobiles/cc/482',
    'sheet_name': 'Mobiles',
    'header': ['Mobile', 'Price (EGP)'],
    'active': True
}, {
    'link': 'https://deals.souq.com/eg-en/health-and-beauty/cc/428',
    'sheet_name': 'Health & Beauty',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/mobiles-and-electronics/cc/424',
    'sheet_name': 'Electronics',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/moms-and-babies/cc/351',
    'sheet_name': 'Moms & Babies',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/home-and-kitchen/cc/493',
    'sheet_name': 'Home & Kitchen',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/stationery/cc/402',
    'sheet_name': 'Office Products',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/car-accessories/cc/400',
    'sheet_name': 'Automotive',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/appliances/cc/348',
    'sheet_name': 'Appliances',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/sports/cc/403',
    'sheet_name': 'Sports',
    'header': ['Name', 'Price (EGP)'],
}, {
    'link': 'https://deals.souq.com/eg-en/toys/cc/495',
    'sheet_name': 'Toys',
    'header': ['Name', 'Price (EGP)'],
}]


def crawl(link):
    page = requests.get(link)
    page_soup = BeautifulSoup(page.text, 'lxml')
    page_items = [
        i.getText().strip() for i in page_soup.select("span.itemTitle a")
    ]
    page_prices = [
        re.search("(^[\d,.]*)", i.getText()).group(0)
        for i in page_soup.select("h5.price span.is")
    ]
    page_pairs = list(zip(page_items, page_prices))
    return page_pairs


def write_sheet(wb, sheet_name, header, pairs, active=False):
    ws = None
    if active:
        ws = wb.active
        ws.title = sheet_name
    else:
        ws = wb.create_sheet(sheet_name)
    ws.cell(1, 1, header[0])
    ws.cell(1, 2, header[1])

    for row in range(2, len(pairs) + 1):
        for col in range(1, 3):
            _ = ws.cell(column=col, row=row, value=f"{pairs[row-2][col-1]}")


wb = Workbook()

for page in webpages:
    print("Getting " + page['sheet_name'])

    content = crawl(page['link'])

    write_sheet(
        wb,
        page['sheet_name'],
        page['header'],
        content,
        active=page.get('active') or False)

wb.save(filename="data.xlsx")
