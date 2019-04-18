# This program scrapes the top 20 most traded stocks on the stock market from Nasdaq.com
# It will further be implemented to be exported or optimized if possible

from bs4 import BeautifulSoup
from urllib.request import urlopen
from datetime import datetime
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy as xl_copy

class NasdaqStockTracker:
    # Name and Price arrays
    prices = []
    names = []
    time = datetime.now()

    def __init__(self):
        # Link of the page that we are scraping data from
        link = 'https://www.nasdaq.com/markets/most-active.aspx'
        # Use the urllib to get URL
        query = urlopen(link)
        # Use bs to put data through an html parser
        soup = BeautifulSoup(query, 'html.parser')

        ###
        # Top 20 stock names that are most active
        ###
        # Table initializer, not needed
        first = soup.find('h3')
        # Append all the names to names array
        for i in range(1, 21):
            first = first.find_next('h3')
            self.names.append(str(i) + ": " + first.get_text())
        ###

        ###
        # Top 20 stock prices of the 20 active names
        ###
        # Initial pointer for table data access
        ptr = soup.find("div", {"class": "genTable"}).find_next('tr').find_next('tr')
        # Append all prices to prices array
        for i in range(0, 20):
            self.prices.append(ptr.find_next('td').find_next('td').find_next('td').find_next('td').getText())
            ptr = ptr.find_next_sibling()
        ###

    # Retrieves stock prices from nasdaq.com
    def get_stock_prices(self):
        print("\nTop 20 Active Stocks and Prices on " + str(self.time.strftime("%c") + " according to NASDAQ"))
        # Print all the names and prices in array
        i = 0
        for name, price in zip(self.names, self.prices):
            i += 1
            if i < 10:
                print('0' + name + ": " + str(price))
            else:
                print(name + ": " + str(price))

    # Writing to a workbook that does not exist yet
    def create_excel(self, ws_name):
        wb = Workbook()
        ws = wb.add_sheet("page")

        # Starts from (1,0), then shifts down by column, then shifts row
        i = 2
        ws.write(1, 0, "Stock Name")
        ws.write(1, 1, "Stock Price")
        for name, price in zip(self.names, self.prices):
            ws.write(i, 0, name)
            ws.write(i, 1, price)
            i += 1

        wb.save(ws_name)

    # Writes the data collected from the stock price function to a excel doc
    def write_to_excel(self):
        wb = open_workbook("Pricing.xls")
        # Create new wb with same data
        wb_copy = xl_copy(wb)
        # Add new sheet on new wb
        ws = wb_copy.add_sheet(self.time.strftime("%m" + "." + "%d" + "." + "%Y"))
        # Starts from (1,0), then shifts down by column, then shifts row
        i = 2
        ws.write(1, 0, "Stock Name")
        ws.write(1, 1, "Stock Price")
        for name, price in zip(self.names, self.prices):
            ws.write(i, 0, name)
            ws.write(i, 1, price)
            i += 1

        wb_copy.save('Pricing.xls')

test = NasdaqStockTracker()

test.write_to_excel()

