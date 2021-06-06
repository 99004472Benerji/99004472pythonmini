
import sys
import openpyxl

new = "python.xlsx"
obj = openpyxl.load_workbook(new)
sud = obj.sheetnames
sh1 = obj['Marks']
sh2 = obj['Hobbies']
print("select the ps number ")
num = int(input())
row = sh1.max_row
column = sh1.max_column
A = 0
B = 0
if num:
    for i in range(1, row + 1):
        for j in range(1, column + 1):
            if sh2.cell(i, j).value == num:
                A = i + 1
                B = j
if A > 0 or B > 0:
    print("data  found")
else:
    print("no data")
    sys.exit()
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Total_data"
print("the data will print on new file ")
workbook.save("new.xlsx")