#coding=utf-8

import openpyxl   #讀寫excel套件

# 讀取Excel
workbook = openpyxl.load_workbook("C:\\Users\\jiunlin\Desktop\\test2.xlsx")

# 取得所有工作表
worksheets = workbook.sheetnames

# 取得第一個工作表
sheet1 = workbook[worksheets[0]]
i = 1
# 取得第二個工作表
sheet2 = workbook[worksheets[1]]
i = 1
# 所有的row loop
for row in sheet1.rows:
    #標題跳過
    if i == 1:
        i = i + 1
        continue

    #判斷每筆row 的 第一個 cell是否為空 是的話跳過
    if row[0].value is None:
        i = i + 1
        continue

    #要比對的值
    array = ["17","19","20","22","24","27","42","55"]

    #將比對的值 loop 去跟 excel的值比對
    for item in array:
        #如果有比對的到會大於-1
        if row[0].value.find(item) > -1:
            #將值給印出來
            print(row[0].value)
            workbook.save('C:\\Users\\jiunlin\Desktop\\test3.xlsx')





