#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter

xl = pd.ExcelFile("hw1_text.xlsx")
categories = ["銀行", "信用卡", "匯率", "台積電", "台灣", "日本"]
columnTitle = ["編號" , "類別", "時間", "標題", "內容"]

sheets = {"銀行":pd.DataFrame(columns = columnTitle), "信用卡":pd.DataFrame(columns = columnTitle), "匯率":pd.DataFrame(columns = columnTitle), 
"台積電":pd.DataFrame(columns = columnTitle), "台灣":pd.DataFrame(columns = columnTitle), "日本":pd.DataFrame(columns = columnTitle)}

def DividetoCategory(news):
	for category in categories:
		title = news["標題"]
		content = news["內容"]
		if(category in title or category in content):
			rowIndex = sheets[category].shape[0]
			sheets[category].loc[rowIndex] = news

def GetWordBySheet():
	df = xl.parse("all")
	nrows = df.shape[0]
	for lineIndex in range(0, nrows):
		DividetoCategory(df.loc[lineIndex,:])

GetWordBySheet()
with pd.ExcelWriter('SortWithGroup.xlsx', mode='a+') as writer:
	for category in categories:
		sheets[category].to_excel(writer, sheet_name= category, index = False, engine='xlsxwriter')