import json
import csv
import re

import lxml
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def crawl(link):
    try:
        page = requests.get(link)
    except requests.exceptions.RequestException:
        print("Couldn't load the webpage.")
        return False

    page_soup = BeautifulSoup(page.text, 'lxml')

    # Get items titles
    page_items = [
        i.getText().strip() for i in page_soup.select("span.itemTitle a")
    ]

    # Get items prices
    # Use regex to extract numbers only
    page_prices = [
        re.search("(^[\d,.]*)", i.getText()).group(0)
        for i in page_soup.select("h5.price span.is")
    ]

    # Zip items and prices in a list and return it
    page_pairs = list(zip(page_items, page_prices))
    return page_pairs


def write_sheet(wb, sheet_name, header, pairs, active=False):
    """ Write an Excel Sheet """
    ws = None

    if active:
        # If it's the first sheet to write
        ws = wb.active
        ws.title = sheet_name
    else:
        ws = wb.create_sheet(sheet_name)

    # Write headers
    ws.cell(1, 1, header[0])
    ws.cell(1, 2, header[1])

    # Write data
    for row in range(2, len(pairs) + 1):
        for col in range(1, 3):
            _ = ws.cell(column=col, row=row, value=f"{pairs[row-2][col-1]}")


def main():
    excel_filename = 'data.xlsx'

    # Load pages.json
    with open('pages.json') as js_file:
        webpages = json.load(js_file)

    # Create new Excel file
    wb = Workbook()

    # Loop through webpages, get each one's data, and write it to the excel sheet
    for page in webpages:
        print("Getting {}... ".format(page['sheet_name']), end='', flush=True)
        content = crawl(page['link'])
        if content:
            print("Done.")
        else:
            continue

        write_sheet(
            wb,
            page['sheet_name'],
            page['header'],
            content,
            active=page.get('active') or False)

    # Save the final excel file
    wb.save(filename=excel_filename)

    print("Data saved to '{}'".format(excel_filename))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        exit(0)