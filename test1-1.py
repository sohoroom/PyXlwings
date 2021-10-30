import xlwings as xw
print ('使用xw.Range()讀寫現正使用中之工作簿的工作表儲存格。')
xw.Range('A1').value = 'xw.Range_A1'
xw.Range('A2').value = 'xw.Range_A2'
xw.Range('A3').value = 'xw.Range_A3'
print(xw.Range('A1').value)
print(xw.Range('A2').value)
print(xw.Range('A3').value)
#針對目前使用中的工作表讀取或寫入時，你可以直接使用 xw.Range() ，不需要另外指定 book 物件或 sheet物件。