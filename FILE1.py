import openpyxl as op
from copy import copy
import re
wb= op.load_workbook(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')
Outbound=wb.get_sheet_by_name('Outbound')
Index=wb.get_sheet_by_name('Index')
Vendor_sub=[]

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
    Fid=''
    Vendor=['DMT','INTERAC','TSYS','BMO','TELUS']
    with open(r'C:\Users\RANA\Desktop\PFA\3 New\3 New\{}\MIM\SQL\insert-JOB-DEV-SYS-INT-PRD.sql'.format(job),'r') as obj:
        for i in obj:
            if re.search(",\d\d\d\d\);",i):
                Fid=i.strip()
            for v in Vendor:
                if v in i :
                    global Vendor_sub
                    Vendor_sub=i.strip().replace(',\'','').replace('\'','').split(' ')
                    break
    Fid=Fid.replace(',','F'+Fid[4])
    Fid=Fid.replace('0);','')
    return Fid

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
    global Vendor_sub
    lst=[]
    for r in range(1,Index.max_row+1):
            if Index.cell(row=r,column=1).value == j:
                lst.append(r)#this will give row names for multiple entries

    if lst:
        Index.cell(row=lst[0],column=3).value=PC_FILEID(j.rstrip())
        Index.cell(row=lst[0],column=4).value=Vendor_sub[0]
        Index.cell(row=lst[0],column=2).value=Vendor_sub[1]+"BOUND"
        lst.pop(0)#eliminate 1st occurance from consideration
        for d in lst[::-1]:  #delete all other rows
            Index.delete_rows(d,1)

    else:print("No entry found for: "+j)

with open(r"C:\Users\RANA\Desktop\PFA\R1 Jobs1.txt","r") as obj:
    for j in obj:
        edit_Outbound(j.rstrip())
        edit_index(j.rstrip())

wb.save(r'C:\Users\RANA\Desktop\PFA\Quantum-II-MIM-IA_Final.xlsx')
