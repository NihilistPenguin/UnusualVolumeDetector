import yfinance as yf
import sys
import os
import numpy as np
import xlwt
from xlwt import Workbook
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import datetime, date
from glob import glob

def getData(ticker):
    sys.stdout = open(os.devnull, "w")
    data = yf.download(ticker, period="1d")
    sys.stdout = sys.__stdout__
    return data[["Close"]]

def closePrice(ticker):
	closePrice = round(np.mean(getData(ticker)['Close']), 2)
	return closePrice

# print(closePrice("WKHS"))

def writeToXls(day, filename):
	with open(filename , "r") as f:
		tickers = f.read().splitlines()

	if day == 0:
		wb = Workbook()
		sheet1 = wb.add_sheet("Sheet 1")
	else: 
		wb = copy(open_workbook('priceTracker.xls'))
		sheet1 = wb.get_sheet(0)

	for i,ticker in enumerate(tickers):
		if day == 0:
			sheet1.write(i, 0, ticker)
		sheet1.write(i, day+1, float(closePrice(ticker)))

	f.close()
	wb.save('priceTracker.xls')


#writeToXls(1)

results_file = glob('results*.txt')
filename = results_file[0]
file_date = filename[-14:-4].split('-')
print(file_date)
initialDate = date(int(file_date[0]), int(file_date[1]), int(file_date[2]))

currentDate = datetime.date(datetime.now())
delta = currentDate - initialDate
writeToXls(delta.days, filename)

# firstDay = input("First day? [y/N]: ")
# if firstDay.lower() == 'y':
# 	# initialDate = datetime.date(datetime.now())
# 	writeToXls(0)
# else:
# 	whatDate = input("What date are you tracking prices from? [yyyy-mm-dd]: ")
# 	tmp = whatDate.split('-')
# 	initialDate = date(int(tmp[0]), int(tmp[1]), int(tmp[2]))
# 	currentDate = datetime.date(datetime.now())
# 	delta = currentDate - initialDate
# 	writeToXls(delta.days)