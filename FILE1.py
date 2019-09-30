import openpyxl as op
from copy import copy
import re
wb= op.load_workbook(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')
Outbound=wb.get_sheet_by_name('Outbound')
Index=wb.get_sheet_by_name('Index')

def rep2ndrow(r,c,job):
    with open(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\insert-JOB-DEV-SYS-INT-PRD.sql'.format(job),'r') as obj:
        for i in obj:
            if "??????????????" in i:
                break
    file=i.strip()
    file=file.replace("'",'')
    file=file.replace(",",'')
    Outbound.cell(row=r,column=c+1).value=file

def PC_FILEID(job):
    with open(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\insert-JOB-DEV-SYS-INT-PRD.sql'.format(job),'r') as obj:
        for i in obj:
            if re.search(",\d\d\d\d\);",i):
                break
    id=i.strip()
    id=id.replace(',','F'+id[4])
    id=id.replace('0);','')
    return id

def edit_Outbound(j):
    mcol=Outbound.max_column
    for i in range(1,Outbound.max_row+1):
        Outbound.cell(row=i,column=mcol+1).value=Outbound.cell(row=i,column=mcol).value
        Outbound.cell(row=i,column=mcol+1)._style=copy(Outbound.cell(row=i,column=mcol)._style)
        if i==2:
            rep2ndrow(i,mcol,j.rstrip())
        elif i==4:
            Outbound.cell(row=i,column=mcol+1).value=j

def edit_index(j):
    lst=[]
    for r in range(1,Index.max_row+1):
            if Index.cell(row=r,column=1).value == j:
                lst.append(r)#this will give row names for multiple entries
    if len(lst)>1:#incase there are more than one rows considering atleast one row fulfills the criteria
        for x in lst:
            if Index.cell(row=x,column=3).value==PC_FILEID(j.rstrip()):
                lst.remove(x)#eliminate correct row from consideration
        for w in lst[::-1]:
            Index.delete_rows(w,1) #delete all other rows from last
    elif len(lst)==1:           #incase only one entry found
        Index.cell(row=lst[0],column=3).value=PC_FILEID(j.rstrip())

with open(r"C:\Users\RANA\Desktop\PFA\R1 Jobs1.txt","r") as obj:
    for j in obj:
        edit_Outbound(j)
        edit_index(j)

wb.save(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')
