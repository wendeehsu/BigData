#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math, os

fields = ["銀行", "信用卡", "匯率", "台積電", "台灣", "日本"]
pointers = ["word", "TF", "DF", "TF-IDF", "全部TF", "全部DF", "全部TF-IDF", "TF期望值", "DF期望值",
			"TF卡方值", "DF卡方值", "MI(用DF)", "Lift(用DF)"]
totalNewsNum = 90000

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
	# condition1
	if(str_a in str_b):
		return str_b
	elif(str_b in str_a):
		return str_a
  
	# condition2
	result=""
	dup=lcs(str_a,str_b)

	# s是"相同部分"的開始位置,e是"相同部分"的結束位置
	s1=str_a.find(dup)
	s2=str_b.find(dup)
	e1=s1+len(dup)
	e2=s2+len(dup)

	# different order and length condition
	if(s1 > s2 and (e1 == len(str_a))):
		result=str_a+str_b[s2+len(dup):len(str_b)]
	elif(s2 > s1 and (e2 == len(str_b))):
		result=str_b+str_a[s1+len(dup):len(str_a)]

	return result

def DFisClose(num1, num2):
	diff = float(abs(num1 - num2))
	if(diff/max(num1, num2) <= inaccuracy):
		return True

	return False

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

def MergeSheet(df_1,df_2):
	df_1 = df_1.append(df_2, ignore_index = True)
	# print(df_1)
	tempList = df_1
	for i in range(df_1.shape[0]):
		for j in range(i+1,df_1.shape[0]):
			key1 = df_1.loc[i, "word"]
			key2 = df_1.loc[j, "word"]
			if(key1 == key2):
				data = {"word": key1,
						"TF": df_1.loc[i, "TF"] + df_1.loc[j, "TF"],
						"DF": df_1.loc[i, "DF"] + df_1.loc[j, "DF"]}
				data["全部TF"] = df_1.loc[i, "全部TF"]
				data["全部DF"] = df_1.loc[i, "全部DF"]
				data["TF-IDF"] = GetTFIDF(data["TF"], data["DF"], df_1.shape[0])
				data["全部TF-IDF"] = GetTFIDF(df_1.loc[i, "全部TF"], df_1.loc[i, "全部DF"], totalNewsNum)
				data["TF期望值"] = GetExpectedValue(df_1.loc[i, "全部TF"], df_1.shape[0])
				data["DF期望值"] = GetExpectedValue(df_1.loc[i, "全部DF"], df_1.shape[0])
				data["TF卡方值"] = GetChiSquare(data["TF"], df_1.loc[i, "全部TF"])
				data["DF卡方值"] = GetChiSquare(data["DF"], df_1.loc[i, "全部DF"])
				data["MI(用DF)"] = GetMI(data["DF"], df_1.loc[i, "全部DF"], df_1.shape[0])
				data["Lift(用DF)"] = GetLift(data["DF"], df_1.loc[i, "全部DF"], df_1.shape[0])
				print(data)
				tempList = tempList[tempList.word != key1].reset_index(drop=True)
				tempList = tempList.append(data, ignore_index = True)

			elif(merge(key1,key2) != "" and DFisClose(df_1.loc[i, "TF"], df_1.loc[j, "TF"])):
				data = {"word": merge(key1,key2),
						"TF": min(df_1.loc[i, "TF"],df_1.loc[j, "TF"]),
						"DF": min(df_1.loc[i, "DF"],df_1.loc[j, "DF"])}
				data["全部TF"] = min(df_1.loc[i, "全部TF"],df_1.loc[j, "全部TF"])
				data["全部DF"] = min(df_1.loc[i, "全部DF"],df_1.loc[j, "全部DF"])
				data["TF-IDF"] = GetTFIDF(data["TF"], data["DF"], df_1.shape[0])
				data["全部TF-IDF"] = GetTFIDF(data["全部TF"], data["全部DF"], totalNewsNum)
				data["TF期望值"] = GetExpectedValue(data["全部TF"], df_1.shape[0])
				data["DF期望值"] = GetExpectedValue(data["全部DF"], df_1.shape[0])
				data["TF卡方值"] = GetChiSquare(data["TF"], data["全部TF"])
				data["DF卡方值"] = GetChiSquare(data["DF"], data["全部DF"])
				data["MI(用DF)"] = GetMI(data["DF"], data["全部DF"], df_1.shape[0])
				data["Lift(用DF)"] = GetLift(data["DF"], data["全部DF"], df_1.shape[0])
				print(data)
				tempList = tempList[tempList.word != key1].reset_index(drop=True)
				tempList = tempList[tempList.word != key2].reset_index(drop=True)
				tempList = tempList.append(data, ignore_index = True)

	return tempList.sort_values("TF", ascending=False).reset_index(drop=True)


file1 = str(input("firstFile name: "))
file2 = str(input("secondFile name: "))
if(os.path.isfile(file1) and file2,os.path.isfile(file2)):
	xl_1 = pd.ExcelFile(file1)
	xl_2 = pd.ExcelFile(file2)
	with pd.ExcelWriter("new"+file1, mode='a+') as writer:
		for field in fields:
			df_1 = xl_1.parse(field)
			df_2 = xl_2.parse(field)
			MergeSheet(df_1,df_2).to_excel(writer, sheet_name= field, index = False, engine = 'xlsxwriter')

