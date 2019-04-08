import PyPDF2
from datetime import datetime
import math 
import openpyxl


DIAS = [
    'SEG',
    'TER',
    'QUA'
    'QUI'
    'SEG'
]
def extract_info(page,information):
    cont=0
    i=0
    j=0
    for letter in page:
        cont+=1
        if(letter=='D'): 
            if(page[cont+20:(cont+24)]=="2019"):
                information[0]=page[(cont+14):(cont+24)]
            if(page[cont-1:(cont+5)]=="Data I"):
                information[4]=page[(cont+27):(cont+42)]
                information[5]=page[(cont+68):(cont+85)]

        if(letter=='A'):            
            if(page[cont-1:(cont+7)]=="Agência:"):
                i=9
                while page[cont+i]!='\n':
                    i+=1
                information[2]=page[cont+9:(cont+i)]
        if(letter=='C'): 
            if(page[cont-1:(cont+17)]=="Causa Fundamental:"):
                cont+=30
                if(page[(cont-11):(cont+9)]=="Falha na comunicação"):
                    information[3]="Embratel"
                else:
                    information[3]=page[cont-11:(cont+14)]
FILE_PATH = '../bin/G3.pdf'
wb = openpyxl.Workbook()#openpyxl.load_workbook('exemplo.xlsx')
ws = wb.active

with open(FILE_PATH, mode='rb') as f:
    reader = PyPDF2.PdfFileReader(f)
    num_pag=0
    information=['aaaa']*9
    while 1:
        page = reader.getPage(num_pag)
        text = page.extractText()
        page = reader.getPage(num_pag+1)
        text += page.extractText()
        extract_info(text,information)
        information[0] = datetime.strptime(information[0], '%d/%m/%Y')
        information[0] = information[0].strftime('%d/%m/%Y')
        datein=datetime.strptime(information[4], '%d/%m/%Y%H:%M')
        information[4] = datetime.strptime(information[4], '%d/%m/%Y%H:%M')
        information[4] = information[4].strftime('%d/%m/%Y %H:%M')
        dateout=datetime.strptime(information[5], '%d/%m/%Y\n %H:%M')
        information[5] = datetime.strptime(information[5], '%d/%m/%Y\n %H:%M')
        information[5] = information[5].strftime('%d/%m/%Y %H:%M')
        datadiff = (dateout-datein)
        #information[1]=(datadiff.seconds)*3600
        information[1]=round((datadiff.seconds)/3600,2)
        ws.append([information[0],information[1],
        information[2],information[3],
        information[4],information[5],
        information[6],information[7],
        information[8]])
        print(information)
        if num_pag+2>=reader.getNumPages():
            break
        num_pag+=2
        
    #print(page.extractText())
    wb.save("exemplo.xlsx")

    
    #print(page.extractText()[5])
