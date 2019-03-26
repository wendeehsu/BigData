#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math
from openpyxl import load_workbook

# Get all news.
xl_news = pd.ExcelFile("../hw1/hw1_text.xlsx")
df_news = xl_news.parse("all")

# Get keywords.
xl_words = pd.ExcelFile("hw2_table.xlsx")
df_words = xl_words.parse("L2_foxconn_keyword")
terms = df_words["term"].tolist()
outputFile = "KeyWords.xlsx"
foxconnNum = 2081
totalNewsNum = 90507

"""
Check if given keyword exist in given news.
0: if keyword does not exist in news.
1: if keyword exists in news (not foxconn).
2: if keyword exists in foxconn news.
"""
def CountWordDF(row, term):
	news = row["標題"] + row["內容"]
	if(term in news):
		if("鴻海" in news or "郭台銘" in news):
			return 2
		return 1
	return 0

# Check all keyword's existance in all news.
def GetDFof(term):
	df_news[term] = df_news.apply(lambda x: CountWordDF(x,term), axis=1)

# Count all keyword's df.
def GetAllDF():
	foxconnDF = list(map(lambda term: df_news.loc[df_news[term] == 2].shape[0] ,terms))
	allDF = list(map(lambda term: df_news.loc[df_news[term] != 0].shape[0] ,terms))
	wordFrame = {"term": terms, "DF" : foxconnDF, "全部DF": allDF}
	
	return pd.DataFrame(wordFrame)

def Jaccard(term):
	upper = term["DF"]
	lower = foxconnNum + term["全部DF"] - term["DF"]
	if(lower == 0):
		return None
	return float(upper)/lower

def MI(term):
	if(term["全部DF"] == 0 or term["DF"] == 0):
		return None
	termMI = term["DF"]/(foxconnNum * term["全部DF"])
	return math.log10(termMI)

def Lift(term):
	upper = float(term["DF"])/foxconnNum
	lower = float(term["全部DF"])/totalNewsNum
	if(lower == 0):
		return None
	return upper/lower

def AddPointer(data):
	data["Jaccard"] = data.apply(Jaccard, axis = 1)
	data["MI"] = data.apply(MI, axis = 1)
	data["Lift"] = data.apply(Lift, axis = 1)

	return data

# Sort data frame by given pointer, displaying top 30 ones.
# value = Lift, MI, Support, Jaccard, or other column name.
def SortDataBy(data, value):
	columnName = value
	if(value == "Support"):
		columnName = "DF"
	selectedFields = {"term" : data["term"], value: data[columnName]}
	sortedDF = pd.DataFrame(selectedFields).sort_values(value, ascending=False).reset_index(drop=True)
	
	return sortedDF.head(30)

# Put the compared (sorted) result of each pointer into a Data Frame.
def CompareData():
	list(map(GetDFof,terms))	# Setup df_news.
	data = AddPointer(GetAllDF())
	output = [SortDataBy(data, "Lift"), SortDataBy(data, "MI"),
			  SortDataBy(data, "Support"), SortDataBy(data, "Jaccard")]
	
	return pd.concat(output, axis = 1)

# Write result to excel.
with pd.ExcelWriter(outputFile, mode='a+') as writer:
	CompareData().to_excel(writer, index = False, engine = 'xlsxwriter')

