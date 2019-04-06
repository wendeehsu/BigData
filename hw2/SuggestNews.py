#!/usr/bin/python
# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import xlsxwriter
import math
from openpyxl import load_workbook

# Get foxconn news.
xl_news = pd.ExcelFile("../hw1/hw1_text.xlsx")
df_news = xl_news.parse("foxconn")

# Get query.
xl_query = pd.ExcelFile("hw2_table.xlsx")
df_query = xl_query.parse("L2_query")

# Get keywords.
df_terms = xl_query.parse("L2_foxconn_keyword")


terms = list(df_terms['term']) # 245筆
query = list(df_query['標題'] + df_query['內容']) # 6筆
data = list(df_news['標題'] + df_news['內容']) #2081筆

del df_terms, df_query

# word2id & id2word.
word2id = {} # (key,value) i.e. 'INCJ': 0
id2word = [] # terms
for i, term in enumerate(terms):
    word2id[term] = i
    id2word.append(term)

del terms

# 算語料集和查詢集的tf值

query_tf = [] # tf of each term for each query
for q in query:
	tmp = []
	for term in id2word:
		tmp.append(q.count(term))
	query_tf.append(tmp)

data_tf = [] # tf of each term for each data
for d in data:
	tmp = []
	for term in id2word:
		tmp.append(d.count(term))
	data_tf.append(tmp)

# 算查詢集裡每個字的df值
df = [] # df of each term
for term in id2word:
	df.append(0)
	for d in data:
		if term in d:
			df[word2id[term]] += 1

# idf
def idf(N, df):
	return math.log(N / df, 10)

# 利用tf-idf建構向量
query_vec = []
data_vec = []
doc_num = len(data)

for tf in data_tf:
	tmp = []
	for i in range(len(id2word)):
		tmp.append(tf[i] * idf(doc_num, df[i]))
	data_vec.append(tmp)

for tf in query_tf:
	tmp = []
	for i in range(len(id2word)):
		tmp.append(tf[i] * idf(doc_num, df[i]))
	query_vec.append(tmp)

# 計算相似度 cosine similarity
def cal_score(q_vec, d_vec):
	q_vec = np.array(q_vec)
	d_vec = np.array(d_vec)
	norm_q = np.sqrt(np.sum(q_vec ** 2))
	norm_d = np.sqrt(np.sum(d_vec ** 2))
	return np.dot(q_vec, d_vec) / (norm_q * norm_d)

# 拿查詢集第一筆資料來對語料集算相似度
q_id = 0
data_score = []
for i, d in enumerate(data_vec):
	data_score.append((i, cal_score(query_vec[q_id], d)))


# 排序取最高3名
data_score.sort(key = lambda item: item[1] if not math.isnan(item[1]) else 0 ,reverse = True)

for d in data_score[:3]:
	print(df_news.iloc[d[0]])