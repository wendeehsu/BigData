#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math
from openpyxl import load_workbook

def CountTF(row, nameList):
	string = str(row['title']) + str(row['content'])
	tf = 0
	for name in nameList:
		tf += string.count(name)

	return tf


def CountCompanyDF(nameList):
	"""
	input: nameList 為公司相關的詞（ex. Company["台泥"]）
	return: relatedNews 表示和給定公司有關的文章
	"""
	csvs = ["bbs.csv", "forum.csv", "news.csv"]
	relatedNews = pd.DataFrame(columns = ['post_time', 'title', 'content'])
	for csv in csvs:
		rawData = pd.read_csv(csv,encoding = 'utf-8')
		data = rawData[['post_time', 'title', 'content']].copy()
		del rawData

		data["TF"] = data.apply(lambda x: CountTF(x,nameList), axis = 1)
		relatedNews = relatedNews.append(data.loc[data["TF"] != 0], ignore_index = True)
		del data

	return relatedNews

Company = {
    "台積電":["台積電", "台gg", "台GG", "GG", "張忠謀", "台積", "魏哲家", "劉德音", "TSMC", "2330"],
    "鴻海":["鴻海", "海邊", "郭台銘", "海公公", "血海", "跑步", "富士康", "foxconn", "2317"],
    "台塑":["台塑", "六輕", "王文淵", "三寶", "四寶", "1301"],
    "南亞":["南亞", "南亞塑膠", "1303"],
    "中華電":["中華電信", "種花", "中華電", "種花電信", "2412"],
    "台化":["台化", "台灣化學纖維", "六輕", "1326"],
    "聯發科":["聯發科", "發哥", "大M", "2454"],
    "中鋼":["中鋼", "中國鋼鐵", "中肛", "肛肛", "林弘男","2002"],
    "統一":["統一", "統二", "桶二", "1216"],
    "宏達電":["宏達電子", "宏達電", "HTC", "hTC", "紅茶", "紅茶店", "王雪紅", "hTㄈ", "HTㄈ", "2498"],
    "台達電":["台達電", "鄭崇華", "2308"],
    "國泰金":["國泰金", "大樹", "花椰菜", "2882"]
}

outputFile = "CompanyNews.xlsx"
with pd.ExcelWriter(outputFile, mode='a+') as writer:
	for company in list(Company.keys()):
		CountCompanyDF(Company[company]).to_excel(writer, sheet_name = company, index = False, engine = 'xlsxwriter')
