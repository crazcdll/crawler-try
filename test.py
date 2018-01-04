import os
from openpyxl import Workbook

# 存储目录
o_path = os.getcwd()
file_path = o_path + '/data/'

wb = Workbook()
ws = wb.active
ws['A1'] = 42
ws.append([1, 2, 3])
import datetime
ws['A2'] = datetime.datetime.now()

# Save the file
wb.save(file_path + "/sample.xlsx")


20140324
20140808
20140822