import random
import openpyxl
import pandas as pd
import numpy as np
import os
from pathlib import Path
import touch
import xlsxwriter

path = 'English_B1_Words.xlsx'
df = pd.read_excel(path)                # read Excel file

words_list = df['WORDS'].tolist()       # dataframe to list
todays_list = random.choices(words_list, k=5)  # random choice from list

if words_list:                          # deleting previously selected words
    print("Today's Words: ", todays_list)
    for word in todays_list:
        words_list.pop(words_list.index(word))
else:
    print("You learned all the words ;)")

row = 1
column = 0
if os.path.isfile(path):
    os.remove(path)                       # existing file was deleted
    workbook = xlsxwriter.Workbook(path)  # creating a new empty file
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', "WORDS")        # added "Words" title

    for words in words_list:              # write to excel
        worksheet.write(row, column, words)
        row += 1
    workbook.close()
    print("List updated!!")
