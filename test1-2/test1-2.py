import xlwings as xw

#visible是否可见。False表示后台运行。 add_book 是否新建一个工作簿
app = xw.App(visible=True, add_book=False)

wb = xw.Book()    #新增一個活頁簿
#wb = xw.Books['工作簿名稱']    引用工作簿
#wb = xw.books.active    使用當前的活頁簿
#wb = app.books.active
#wb = app.books.open(r'file_path')

sht = wb.sheets['工作表1']
print(sht.name)
#sht = wb.sheets.active    使用當前的工作表
#sht = wb.sheets[0]    使用第一個工作表
#sht = wb.sheets['Sheet1']    使用名稱的工作表

sht.range('A1').value = 'sht.range_A1'
sht.range('A2').value = 'sht.range_A2'
sht.range('A3').value = 'sht.range_A3'
sht.autofit()

A4 = sht.range('A4')
A4.api.NumberFormat = "@"
sht.range('A5').number_format = '@'    #number_format注意大小寫，與api不同
#sht.range('A6').api.NumberFormat = "General"    應使用"G/通用格式"，而非General
sht.range('A6').number_format = 'G/通用格式'
sht.range('A7').number_format = '0.00'
sht.range('A8').api.NumberFormat = "@"    #@代表文字格式
sht.range('B1:B10').number_format = '@'

sht.range('A4').value = '1-4'
sht.cells(5,'A').value = '1-5'
sht.cells(9, 1).value = '1-9'    #使用cells，先行後列
#sht.range('A5').value = '1-5'
sht.range('A6').value = '3.1415926'
sht.range('A7').value = '3.1415926'
sht.range('A8').value = '3.1415926'
sht.range('B1:B8').value = '1-1'    #B1:B8全部填值為1-1
sht.range('A2:A6').api.Copy(sht.range('A10').api)    #將A2:A6複製至A10起始(向下填值)
sht.range('C1').value = [1,2,3,4]    #預設為同行向右填值，C1起始即為C1:C4
sht.range('C3').expand('table').value = [['a','b','c'],['d','e','f'],['g','h','i']]    #多行列填值，使用多維陣列
# 清除range的內容
#rng.clear_contents()
# 清除格式和內容
#rng.clear()

#wb.name = 'test1-2'    活頁簿名稱需使用save儲存改變，無法直接更改
sht.name = 'sheet1'
print('sht.index =',sht.index)
print('sht.name =',sht.name)
wb.sheets[1].name = 'sheet2'
print('wb.sheets[0].name =', wb.sheets[0].name)
print('wb.sheets[1].name =', wb.sheets[1].name)
print('wb.sheets[2].name =', wb.sheets[2].name)
#sht2 = wb.sheets.add('新建工作表', after=sht)
print(sht.range('A1:A8').options(ndim=1).value)
print(sht.range('A1:B8').options(ndim=2).value)

wb.save('C:\\Users\\user\\Desktop\\xlwingsPY\\test1-2.xlsx')
#wb.save('C:/Users/user/Desktop/xlwingsPY/test1-2.xlsx')    windows儲存目錄可使用\\或/
#print(wb.name)
print(wb.fullname)
#wb.close()