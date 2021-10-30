import xlwings as xw
import twd97
#import tkinter
#from tkinter import messagebox

def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A22").value = "hello world!"

@xw.func
def hello(name):
    return "hello {0}".format(name)

@xw.func
def decorate(name2):
    return str(name2)

@xw.func
def fromwgs84(lat,lon):
    temp = twd97.fromwgs84(lat,lon)
    t1 = float(temp[0])
    t2 = float(temp[1])
    temp2 = (round(t1,6),round(t2,6))
    strtemp = str(temp2)                    #tuple to string
#messagebox.showinfo("Information",temp)    #for debug
    return strtemp.strip('()')              #strip (&)
# don't use "long"

@xw.func
def towgs84(x,y):
    tmp = twd97.towgs84(x,y)
    tt1 = float(tmp[0])
    tt2 = float(tmp[1])
    tmp2 = (round(tt1,6),round(tt2,6))
    tstrtemp = str(tmp2)
    return tstrtemp.strip('()')
#    return tt1
