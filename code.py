import openpyxl
def Excel_info():
    loc = "python.xlsx"
    wb = openpyxl.load_workbook(loc)

    sh = wb["Marks"]
    for n in sh['A']:
        print(n.value)

    a = {}
    file_obj = wb.active
    columns = ['PS NO']

    for sheet in wb.worksheets:
        for i in range(2,sheet.max_row+1):
            for j in range(2,sheet.max_column+1):
                if sheet.cell(row=i,column=1).value not in a:
                    a[sheet.cell(row=i,column=1).value] = []
                if sheet.cell(row=1,column=j).value not in columns:
                    columns.append(sheet.cell(row=1,column=j).value)
                a[sheet.cell(row=i,column=1).value].append(sheet.cell(row=i,column=j).value)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Total Info"

    for i in range(len(columns)):
        sheet.cell(row=1,column=i+1).value = columns[i]
    keys = [int(input('Select a PS NO: '))]
    for i in range(2,len(keys)+2):
        sheet.cell(row = i,column=1).value=keys[i-2]
        c=2
    for j in a[keys[i-2]]:
        sheet.cell(row=i,column=c).value=j
        c+=1
    workbook.save("Output.xlsx")
Excel_info()