import openpyxl as op
from copy import copy
wb= op.load_workbook(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')
sheet=wb.get_sheet_by_name('Inbound')

def rep2ndrow(r,c,job):
    with open(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\insert-JOB-DEV-SYS-INT-PRD.sql'.format(job),'r') as obj:
        for i in obj:
            if "??????????????" in i:
                break
    file=i.strip()
    file=file.replace("'",'')
    file=file.replace(",",'')
    sheet.cell(row=r,column=c+1).value=file

with open(r"C:\Users\RANA\Desktop\PFA\R1 Jobs1.txt","r") as obj:
    for j in obj:
        mcol=sheet.max_column
        for i in range(1,sheet.max_row+1):
            sheet.cell(row=i,column=mcol+1).value=sheet.cell(row=i,column=mcol).value
            sheet.cell(row=i,column=mcol+1)._style=copy(sheet.cell(row=i,column=mcol)._style)
            if i==2:
                rep2ndrow(i,mcol,j.rstrip())
            elif i==4:
                sheet.cell(row=i,column=mcol+1).value=j

wb.save(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')

