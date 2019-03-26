#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math
from openpyxl import load_workbook

xl = pd.ExcelFile("hw2_table.xlsx")
df = xl.parse("L2_foxconn_keyword")
outputFile = "KeyWords.xlsx"

xl_News = pd.ExcelFile("../hw1/hw1_table.xlsx")
df_News_2gram = xl_News.parse("HONHAI_2gram", header = 2)
df_News_3gram = xl_News.parse("HONHAI_3gram", header = 2)
terms = df["term"].tolist()

# Get keywords and their related pointers.
def FetchKeyword():
	keyword = df_News_2gram[df_News_2gram['詞'].isin(terms)].reset_index()
	keyword_3gram = df_News_3gram[df_News_3gram['詞'].isin(terms)].reset_index()
	keywordList = keyword.append(keyword_3gram, ignore_index = True)

	return keywordList.reset_index()

# Get keyword's Jaccard coefficient.
def Jaccard(news):
	upper = news["DF"]
	lower = 2081 + news["全部DF"] - news["DF"]

	return float(upper)/lower

# Sort data frame by given pointer, displaying top 30 ones.
# value = Lift, MI, Support, Jaccard, or other column name.
def SortDataBy(data, value):
	columnName = value
	if(value == "Lift"):
		columnName = "Lift(用DF)"
	elif(value == "MI"):
		columnName = "MI(用DF)"
	elif(value == "Support"):
		columnName = "DF"
	selectedFields = {"詞" : data["詞"], value: data[columnName]}
	sortedDF = pd.DataFrame(selectedFields).sort_values(value, ascending=False).reset_index(drop=True)
	
	return sortedDF.head(30)

# Put the compared (sorted) result of each pointer into a Data Frame.
def CompareData():
	data = FetchKeyword()
	data["Jaccard"] = data.apply(Jaccard, axis=1)
	output = [SortDataBy(data, "Lift"), SortDataBy(data, "MI"),
			  SortDataBy(data, "Support"), SortDataBy(data, "Jaccard")]
	
	return pd.concat(output, axis = 1)

# Write result to excel.
with pd.ExcelWriter(outputFile, mode='a+') as writer:
	CompareData().to_excel(writer, index = False, engine = 'xlsxwriter')
