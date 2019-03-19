#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd

xl = pd.ExcelFile("hw1_table.xlsx")
MI_TF_IDF = "MI_TF_IDF"
Keywords = "Keywords"
sheets = ["ALL_2gram", "ALL_3gram", "INDUSTRY_2gram", "INDUSTRY_3gram", "HONHAI_2gram", "HONHAI_3gram"]
goals = ["TF", "DF", "TF-IDF", "全部TF", "全部DF", "全部TF-IDF", "TF期望值", "DF期望值", "TF卡方值(保留正負號)", "DF卡方值(保留正負號)", "MI(用DF)", "Lift(用DF)", MI_TF_IDF]
outputFile = "out.xlsx"

# variables
visibleRow = 30
inaccuracy = 0.01

def WordSimilar(str1, str2):
    for chr in str1:
        if chr in str2:
            return True

    return False

def DFisClose(num1, num2):
    diff = float(abs(num1 - num2))
    if(diff/max(num1, num2) <= inaccuracy):
        return True

    return False

def CheckSheetBy(sheetName, goal):
    df = xl.parse(sheetName, header=2)
    if(goal == MI_TF_IDF):
        df[MI_TF_IDF] = df["TF-IDF"] * df["MI(用DF)"]
    
    result = df.sort_values(goal, ascending=False)
    keywords = result["詞"].head(visibleRow).tolist()
    DF = result["DF"].head(visibleRow).tolist()
    values = result[goal].head(visibleRow).tolist()
    df1 = {Keywords: keywords, "DF": DF, goal: values}

    return pd.DataFrame(df1)

def MergeBy(field, goal):
    if (field == 0):
        mergeDf = pd.concat([CheckSheetBy(sheets[0],goal), CheckSheetBy(sheets[1],goal)], ignore_index = True)
    elif (field == 1):
        mergeDf = pd.concat([CheckSheetBy(sheets[2],goal), CheckSheetBy(sheets[3],goal)], ignore_index = True)
    else:
        mergeDf = pd.concat([CheckSheetBy(sheets[4],goal), CheckSheetBy(sheets[5],goal)], ignore_index = True)
    
    tempList = mergeDf
    for i in range(len(mergeDf[Keywords])):
        key1 = mergeDf[Keywords][i]
        for j in range(i+1, len(mergeDf[Keywords])):
            key2 = mergeDf[Keywords][j]
            if(WordSimilar(key1, key2)):
                num1 = mergeDf.loc[mergeDf[Keywords] == key1]["DF"].values[0]
                num2 = mergeDf.loc[mergeDf[Keywords] == key2]["DF"].values[0]
                if(DFisClose(num1,num2)):
                    r = key1
                    if(num2 < num1):
                        r = key2
                    tempList = tempList[tempList.Keywords != r]

    return tempList.sort_values(goal, ascending=False).reset_index(drop=True)

with pd.ExcelWriter('out.xlsx', mode='a') as writer:
    for i in [0,1,2]:
        resultDF = pd.DataFrame()
        if(i == 0):
            for goal in goals[:3]:
                resultDF = pd.concat([resultDF, MergeBy(i,goal)], axis = 1)
        else:
            for goal in goals:
                resultDF = pd.concat([resultDF, MergeBy(i,goal)], axis = 1)

        resultDF.to_excel(writer, sheet_name= sheets[2*i+1], index = False)
    