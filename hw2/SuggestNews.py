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

# Get query.
xl_query = pd.ExcelFile("hw2_table.xlsx")
df_query = xl_query.parse("L2_query")

# Get keywords.
df_terms = xl_query.parse("L2_foxconn_keyword")

terms = list(df_terms['term'])
query = list(df_query['標題']+df_query['內容'])
data = list(df_news['標題']+df_news['內容'])

del df_terms, df_query