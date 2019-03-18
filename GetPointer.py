#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math

patchLimit = 10
xl = pd.ExcelFile("hw1_text.xlsx")
df = xl.parse("all")
totalNewsNum = df.shape[0]
outputFile = "KeyWords.xlsx"
fields = ["銀行", "信用卡", "匯率", "台積電", "台灣", "日本"]
fieldNum = {"銀行":2, "信用卡":3, "匯率":5, 
			"台積電":7, "台灣":11, "日本":13}  # match fields with num for labeling
columnTitles = ["word" , "field",
				"all_TF", "all_DF", "bank_TF", "bank_DF",
				"card_TF", "card_DF", "rate_TF", "rate_DF",
				"tsmc_TF", "tsmc_DF", "tw_TF", "tw_DF",
				"japan_TF", "japan_DF"]
pointers = ["word", "TF", "DF", "TF-IDF", "全部TF", "全部DF", "全部TF-IDF", "TF期望值", "DF期望值",
			"TF卡方值", "DF卡方值", "MI(用DF)", "Lift(用DF)"]

commenWords = [] 	# words that should not be keywords
with open('commenWord.txt', encoding = "utf-8") as f:
	commenWords = f.read().splitlines()

def HasNumOrEn(string):
	for ch in string:
		# chinese is between \u4e00 and \u9fff
		if(ch < '\u4e00' or ch > '\u9fff'):
			return True
	return False

def CutContent(string):
	wordList = pd.DataFrame(columns = ["word", "TF"])
	for n in range(2,7):
		for i in range(len(string)-n+1):
			word = string[i:i+n]
			if(not HasNumOrEn(word) and word not in commenWords):
				if(word not in wordList["word"].tolist()):
					newWord = {"word": word, "TF":1}
					wordList = wordList.append(newWord, ignore_index = True)
				else:
					wordList.loc[wordList["word"] == word, "TF"] += 1

	return wordList 	# a frame with each unique term's TF

def CleanSheetList(keyList, lineIndex):
	cleanList = keyList[keyList.all_DF > 1]
	if(lineIndex % 50 == 0):
		cleanList = cleanList[cleanList.all_TF > 3]

	return cleanList

def GetFullColumnData(wordWithTF, belongField):
	fullFrame = pd.DataFrame(columns = columnTitles)
	data = {"word": wordWithTF["word"], "field": belongField, 
			"all_TF": wordWithTF["TF"], "all_DF": 1}
	for fieldIndex in range(len(fields)):
		if(belongField % fieldNum[fields[fieldIndex]] == 0):
			data[columnTitles[2*fieldIndex+4]] = wordWithTF["TF"]
			data[columnTitles[2*fieldIndex+5]] = 1

	fullFrame = fullFrame.append(data, ignore_index = True).fillna(0)
	return fullFrame 		# a full frame of each term's index

def MergeWordTF(titleWordTF, contentWordTF):
	for lineIndex in range(titleWordTF.shape[0]):
		wordinTitle = titleWordTF.loc[lineIndex, "word"]
		if(wordinTitle not in contentWordTF["word"].tolist()):
			contentWordTF = contentWordTF.append(titleWordTF.loc[lineIndex], ignore_index = True)
		else:
			contentWordTF.loc[contentWordTF["word"] == wordinTitle, "TF"] += titleWordTF.loc[lineIndex, "TF"]

	return contentWordTF 	# a frame with unique term's TF from title, content

def GetWordByArticle(news):
	belongField = 1 	 # record the fields the news belongs to 
	title = news["標題"]
	content = news["內容"]
	articleList = pd.DataFrame(columns = columnTitles)  # create empty list to store keywords in article
	keywordPool = MergeWordTF(CutContent(title), CutContent(content))

	for field in fields: 	# label its belonged fields
		if (field in title or field in content):
			belongField *= fieldNum[field]
	
	for lineIndex in range(keywordPool.shape[0]):
		data = GetFullColumnData(keywordPool.loc[lineIndex], belongField)
		articleList = articleList.append(data, ignore_index = True)

	return articleList	# a full frame of pointers for unique words in news

def GetWordBySheet():
	sheetList = pd.DataFrame(columns = columnTitles)  # create empty list to store keywords for each sheet
	for lineIndex in range(totalNewsNum):
		print("lineIndex: ",lineIndex)
		keywords = GetWordByArticle(df.loc[lineIndex])
		for wordIndex in range(keywords.shape[0]):
			keyword = keywords.loc[wordIndex]
			if(keyword["word"] not in sheetList["word"].tolist()):
				sheetList = sheetList.append(keyword, ignore_index = True)
			else:
				for columnIndex in range(1, len(columnTitles)):
					if(columnIndex == 1):
						sheetList.loc[sheetList["word"] == keyword["word"], "field"] *= keyword["field"]
					else:
						sheetList.loc[sheetList["word"] == keyword["word"], columnTitles[columnIndex]] += keyword[columnTitles[columnIndex]]

		if (lineIndex != 0 and (lineIndex % patchLimit == 0)):
			sheetList = CleanSheetList(sheetList, lineIndex)
	
	return sheetList 	# a full dataFrame with all unique word's TF, DF in each field

def GetTFIDF(tf, df, n):
	log_tf = math.log10(tf)
	log_Ndf = math.log10(float(n)/df)
	return (1+log_tf)*log_Ndf

def GetChiSquare(value, expectedValue):
	chi = (value-expectedValue)**2
	chi /= expectedValue
	if(value < expectedValue):
		chi *= -1

	return chi

def GetExpectedValue(value, fieldNewsNum):
	return value * float(fieldNewsNum) / totalNewsNum

def GetMI(wordDF, allDF, fieldNewsNum):
	return math.log10(float(wordDF)/(allDF*fieldNewsNum))

def GetLift(wordDF, allDF, fieldNewsNum):
	upper = float(wordDF)/fieldNewsNum
	lower = float(allDF)/totalNewsNum

	return upper/lower

def GetDataByField(rawData, fieldName):
	fieldIndex = fields.index(fieldName)
	data = rawData[rawData.field % fieldNum[fieldName] == 0] # column = columnTitles
	data = data.reset_index()
	fieldSheet = pd.DataFrame(columns = pointers)
	fieldSheet["word"] = data["word"]
	fieldSheet["TF"] = data[columnTitles[2*fieldIndex+4]]
	fieldSheet["DF"] = data[columnTitles[2*fieldIndex+5]]
	fieldSheet["全部TF"] = data["all_TF"]
	fieldSheet["全部DF"] = data["all_DF"]

	fieldNewsNum = data.shape[0]
	for i in range(fieldNewsNum):
		fieldSheet.loc[i, "TF-IDF"] = GetTFIDF(fieldSheet.loc[i, "TF"], fieldSheet.loc[i, "DF"], fieldNewsNum)
		fieldSheet.loc[i, "全部TF-IDF"] = GetTFIDF(fieldSheet.loc[i, "全部TF"], fieldSheet.loc[i, "全部DF"], totalNewsNum)
		fieldSheet.loc[i, "TF期望值"] = GetExpectedValue(fieldSheet.loc[i, "全部TF"], fieldNewsNum)
		fieldSheet.loc[i, "DF期望值"] = GetExpectedValue(fieldSheet.loc[i, "全部DF"], fieldNewsNum)
		fieldSheet.loc[i, "TF卡方值"] = GetChiSquare(fieldSheet.loc[i, "TF"], fieldSheet.loc[i, "全部TF"])
		fieldSheet.loc[i, "DF卡方值"] = GetChiSquare(fieldSheet.loc[i, "DF"], fieldSheet.loc[i, "全部DF"])
		fieldSheet.loc[i, "MI(用DF)"] = GetMI(fieldSheet.loc[i, "DF"], fieldSheet.loc[i, "全部DF"], fieldNewsNum)
		fieldSheet.loc[i, "Lift(用DF)"] = GetLift(fieldSheet.loc[i, "DF"], fieldSheet.loc[i, "全部DF"], fieldNewsNum)

	return fieldSheet
		
with pd.ExcelWriter(outputFile, mode='a+') as writer:
	for field in fields:
		data = GetDataByField(GetWordBySheet(),field)
		print(field, "\n", data)
		data.to_excel(writer, sheet_name= field, index = False, engine = 'xlsxwriter')
