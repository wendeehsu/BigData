#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math

xl = pd.ExcelFile("SortWithGroup.xlsx")
sheets = ["銀行", "信用卡", "匯率", "台積電", "台灣", "日本"]
columnTitles = ["word" , "TF", "DF"]
outputFile = "KeyWords.xlsx"
patchLimit = 50
inaccuracy = 0.01

commenWords = [] # words that should not be keywords
with open('commenWord.txt', encoding = "utf-8") as f:
	commenWords = f.read().splitlines()

def HasNumOrEn(string):
	for ch in string:
		if(ch < '\u4e00' or ch > '\u9fff'):
			return True
	return False

def DFisClose(num1, num2):
	diff = float(abs(num1 - num2))
	if(diff/max(num1, num2) <= inaccuracy):
		return True

	return False

def CutContent(string):
	wordList = []
	for n in range(2,7):
		for i in range(len(string)-n+1):
			word = string[i:i+n]
			if(not HasNumOrEn(word) and word not in commenWords):
				wordList.append(word)

	return wordList

def CleanSheetList(keyList):
	# expectedTF = keyList["TF"].mean()
	# cleanList = keyList[keyList.TF > expectedTF]
	cleanList = keyList[keyList.TF >= 3]

	return cleanList

def GetWordByArticle(news):
	articleList = pd.DataFrame(columns = columnTitles)  # create empty list to store keywords in article
	title = news["標題"]
	content = news["內容"]
	keywordPool = CutContent(title)
	keywordPool.extend(CutContent(content))

	for word in keywordPool:
		if(word not in articleList["word"].tolist()):  # add word to list
			news = {columnTitles[0]:word, columnTitles[1]:1, columnTitles[2]:1}
			articleList.loc[articleList.shape[0]] = news
		else:  # update `TF`
			articleList.loc[articleList["word"] == word, "TF"] += 1

	return articleList

def GetWordBySheet(sheetName):
	df = xl.parse(sheetName)
	nrows = df.shape[0]
	sheetList = pd.DataFrame(columns = columnTitles)  # create empty list to store keywords for each sheet
	for lineIndex in range(nrows):
		keywords = GetWordByArticle(df.loc[lineIndex])
		for wordIndex in range(keywords.shape[0]):
			keyword = keywords.loc[wordIndex]
			if(keyword["word"] not in sheetList["word"].tolist()):
				sheetList.loc[sheetList.shape[0]] = keyword
			else:
				sheetList.loc[sheetList["word"] == keyword["word"], "TF"] += keyword["TF"]
				sheetList.loc[sheetList["word"] == keyword["word"], "DF"] += keyword["DF"]

		if (lineIndex != 0 and (lineIndex % patchLimit == 0)):
			sheetList = CleanSheetList(sheetList)
		print("len: ", sheetList.shape[0],"\n",sheetList.head())
	return sheetList

def AddPointer(keywordList):
	n = keywordList.shape[0]
	for i in range(n):
		log_tf = math.log10(keywordList.loc[i,"TF"])
		log_Ndf = math.log10(float(n)/keywordList.loc[i,"DF"])
		keywordList.loc[i,"TF-IDF"] = (1+log_tf)*log_Ndf

	return keywordList

def lcs(str_a, str_b):
	if len(str_a) == 0 or len(str_b) == 0:
		return 0
  
	lcs_str=""
	max_len = 0
  
	dp = [0 for _ in range(len(str_b) + 1)]
	for i in range(1, len(str_a) + 1):
		left_up = 0
		for j in range(1, len(str_b) + 1):
			up = dp[j]
			if str_a[i-1] == str_b[j-1]:
				dp[j] = left_up + 1
				max_len = max([max_len, dp[j]])
				if max_len == dp[j]:
					lcs_str = str_a[i-max_len:i]
			else:
				dp[j] = 0
			left_up = up
	return(lcs_str)

# find longest common substring
def merge(str_a, str_b):   
	result=""
	dup=lcs(str_a,str_b)
	s1=str_a.find(dup)
	s2=str_b.find(dup)
	e1=s1+len(dup)
	e2=s2+len(dup)
  
	# different order condition 
	if(s1 > s2 & e1 and len(str_a)):
		result=str_a+str_b[s2+len(dup):len(str_b)]
	if(s2 > s1 & e2 and len(str_b)):
		result=str_b+str_a[s1+len(dup):len(str_a)]
  
	return result

def RemoveDuplicateKeyword(keywordList):
	tempList = keywordList
	for i in range(keywordList.shape[0]):
		for j in range(i+1, keywordList.shape[0]):
			key1 = keywordList.loc[i, "word"]
			key2 = keywordList.loc[j, "word"]
			if(merge(key1, key2) != ""):
				DF1 = keywordList.loc[keywordList["word"] == key1,"DF"].values[0]
				DF2 = keywordList.loc[keywordList["word"] == key2,"DF"].values[0]
				if(DFisClose(DF1,DF2)): # merge two rows
					mergedWord = merge(key1,key2)
					mergedNews = {"word": mergedWord, 
								  "TF": keywordList.loc[keywordList["word"] == key1,"TF"].values[0],
								  "DF": DF1}
					if(DF1 < DF2):
						mergedNews["TF"] = keywordList.loc[keywordList["word"] == key2,"TF"].values[0]
						mergedNews["DF"] = DF2
					
					tempList = tempList[tempList.word != key1].reset_index(drop=True)
					tempList = tempList[tempList.word != key2].reset_index(drop=True)
					tempList = tempList.append(mergedNews, ignore_index = True)
	tempList = tempList[tempList.TF > 3]

	return tempList.sort_values("TF", ascending=False).reset_index(drop=True)


with pd.ExcelWriter(outputFile, mode='a+') as writer:
	# keywordSheet = AddPointer(RemoveDuplicateKeyword(GetWordBySheet(sheets[0])))
	# keywordSheet.to_excel(writer, sheet_name= sheets[0], index = False, engine = 'xlsxwriter')
	for sheet in sheets:
		keywordSheet = AddPointer(RemoveDuplicateKeyword(GetWordBySheet(sheet)))
		keywordSheet.to_excel(writer, sheet_name= sheet, index = False, engine = 'xlsxwriter')
