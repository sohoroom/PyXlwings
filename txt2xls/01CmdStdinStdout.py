import fileinput

# 以迴圈逐行處理
for line in fileinput.input():

    # 去除結尾處換行字元
    line = line.rstrip()

    print(line)
#
#import sys 
#for line in sys.stdin: 
#    print('Output:', line.rstrip())
#