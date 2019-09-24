import openpyxl as op
from copy import copy
wb= op.load_workbook(r'C:\Users\RANA\Desktop\Book1.xlsx')
sheet=wb.get_sheet_by_name('Index')
print(sheet.max_row)
sheet.delete_rows(5)
print(sheet.max_row)
'''
with open(r"C:/Users\RANA\Desktop\PFA\R1 Jobs1.txt","r") as obj:
    for j in obj:
        lst=[]
        for r in range(1,sheet.max_row+1):
            if sheet.cell(row=r,column=1).value == j:
                lst.append(r) #this will give row names for multiple entries
        if len(lst)>1
            for x in lst:
               #if sheet.cell(row=x,column=3).value!=PCFILEID            
'''




