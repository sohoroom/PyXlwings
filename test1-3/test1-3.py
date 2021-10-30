import xlwings as xw


def hello_xlwings():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def hello(name):
    return "hello {0}".format(name)
#自定義函數hello()，以@xw.func開頭

@xw.func
def pysum(a,b):
    return a + b