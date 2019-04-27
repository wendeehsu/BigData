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

