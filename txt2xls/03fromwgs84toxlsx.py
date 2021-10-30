#python fromwgs84.py <input_file.txt>
#

import fileinput
import twd97
import openpyxl
from openpyxl import Workbook

wb = Workbook()
sheet = wb.active
sheet.title = 'fromwgs84'
sheet.column_dimensions['A'].width = 16
sheet.column_dimensions['B'].width = 16
sheet.column_dimensions['C'].width = 16
sheet.column_dimensions['D'].width = 16
#sheet.row_dimensions[1].height = 32
# 以迴圈逐行處理
for line in fileinput.input():

    # 去除結尾處換行字元
    line = line.rstrip()

    #print(line)
    #lat,lng 字串分割
    #split()分隔，預設為空白，可為換行(\n)、tab(\t)、逗號(,)或其他
    #分割後存入串列list[]，不會自動去除多餘的空白，可用strip()處理
    #或使用str.replace(" ","")取代所有的空白
    line = line.replace(" ","")
    latlng = line.split(',')
    lat = float(latlng[0])
    lng = float(latlng[1])
    xy = twd97.fromwgs84(lat,lng)
    latlngxy = list(latlng) + list(xy)
    print(latlngxy)
    i = 0
    sheet.append(latlngxy)
    i = i + 1
wb.save('fromwgs84.xlsx')

#
#import sys 
#for line in sys.stdin: 
#    print('Output:', line.rstrip())
#