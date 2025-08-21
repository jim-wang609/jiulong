import math
import openpyxl
import pandas as pd
import numpy as np


# """将表格列字母（如'A'、'AB'）转换为数字（如1、28）"""
def column_letter_to_number(letter):
    number = 0
    for char in letter.upper():
        number = number * 26 + (ord(char) - ord('A') + 1)
    return number


file_path = '7.xlsx'
sheet_name = '8成品二算料单标准（修改版）'
pf = pd.read_excel(file_path, sheet_name=sheet_name)
wb = openpyxl.load_workbook(file_path)
ws = wb[sheet_name]
print(pf)
arr = np.array(pf)
print(arr)
print(len(arr)-1)
# print(arr[2][column_letter_to_number('K')-1]-arr[2][column_letter_to_number('C')-1]%arr[2][column_letter_to_number('E')-1])
count = 0
for i in range(2, len(arr) - 5):
    num = 0
    if (arr[i][column_letter_to_number('K') - 1] - arr[i][column_letter_to_number('C') - 1]) % arr[i][
        column_letter_to_number('E') - 1] < 24:
        num = math.floor(
            (arr[i][column_letter_to_number('K') - 1] - arr[i][column_letter_to_number('C') - 1]) / arr[i][
                column_letter_to_number('E') - 1] + 1)
    else:
        num = math.ceil(
            (arr[i][column_letter_to_number('K') - 1] - arr[i][column_letter_to_number('C') - 1]) / arr[i][
                column_letter_to_number('E') - 1] + 1)
    if num >= 36:
        ws.insert_rows(i + 3 + count)
        count = count + 1
        print(count)
wb.save('out.xlsx')

