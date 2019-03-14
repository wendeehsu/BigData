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

def HasNumOrEn(string):
    for ch in string:
        if(ch < '\u4e00' or ch > '\u9fff'):
            return True
    return False

def CutContent(string):
    wordList = []
    for n in range(2,7):
        for i in range(len(string)-n+1):
            word = string[i:i+n]
            if(not HasNumOrEn(word)):
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

    return sheetList

def AddPointer(keywordList):
    n = keywordList.shape[0]
    for i in range(n):
        log_tf = math.log10(keywordList.loc[i,"TF"])
        log_Ndf = math.log10(float(n)/keywordList.loc[i,"DF"])
        keywordList.loc[i,"TF-IDF"] = (1+log_tf)*log_Ndf

    return keywordList

with pd.ExcelWriter(outputFile, mode='a+') as writer:
    for sheet in sheets:
        keywordSheet = AddPointer(GetWordBySheet(sheet))
        keywordSheet.to_excel(writer, sheet_name= sheet, index = False, engine = 'xlsxwriter')
