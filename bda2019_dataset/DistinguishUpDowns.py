#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math
import matplotlib.pyplot as plt
from openpyxl import load_workbook

xl = pd.ExcelFile("SortedCompanyStocks.xlsx")
CompanyStocks = ['1301 台塑', '2330 台積電', '2317 鴻海', '1303 南亞', 
				'2412 中華電', '1326 台化', '2454 聯發科', '2002 中鋼', 
				'1216 統一', '2498 宏達電', '2308 台達電', '2882 國泰金'  ]

for stock in CompanyStocks:
	df = xl.parse(stock)
	df["周均線"] = df["收盤價(元)"].rolling(window=5).mean()
	df["月均線"] = df["收盤價(元)"].rolling(window=20).mean()