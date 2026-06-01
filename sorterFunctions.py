from pypdf import PdfReader
import pdfplumber,re
import docx2pdf
import os
from pypdf import PdfWriter
from collections import defaultdict

def findNextNumber(text,index):
    if text[index].isdigit()==True:
        endIndex=index
        while text[endIndex].isdigit()==True:
            endIndex+=1
    while text[index].isdigit()==False:
        index+=1
        endIndex=index
        while text[endIndex].isdigit()==True:
            endIndex+=1
    if endIndex>index:
        return text[index:endIndex]

def orderPDF(pdfPath):
    reader = PdfReader(pdfPath)
    currentOrder = []
    with pdfplumber.open(pdfPath) as pdf:
        for index in range(len(reader.pages)):
            page = pdf.pages[index]
            pageTables = page.extract_tables()
            DC_index=findTablePosition(pageTables,'Loc. #:')
            BOLNum_index=findTablePosition(pageTables,'Bill of Lading Number:')
            orderInfoTable=findTablePosition(pageTables,'Customer Order Number')
            bol_pos=[]
            bol_index=1
            while pageTables[orderInfoTable[0][0]][orderInfoTable[0][1]+bol_index][0] != '' and pageTables[orderInfoTable[0][0]][orderInfoTable[0][1]+bol_index][0] != 'GRAND TOTAL':
                currentRow=pageTables[orderInfoTable[0][0]][orderInfoTable[0][1]+bol_index]
                PO_DC_Num=currentRow[0]
                PONum=PO_DC_Num.split("-")[1].split("-")[0]
                cartons=findNextNumber(currentRow[1],0)
                weight=findNextNumber(currentRow[2],0)
                DC=findNextNumber(DC_index[1],DC_index[1].find('Loc. #:'))
                BOLNum=findNextNumber(BOLNum_index[1],BOLNum_index[1].find('Bill of Lading Number:'))
                bol_pos.append([PO_DC_Num,cartons,weight])
                bol_index+=1
            currentOrder.append([index,DC,BOLNum,bol_pos])

    sortedByDCOrder = sorted(currentOrder, key=lambda row: row[1])

    return sortedByDCOrder

def findTablePosition(tableArr,target):
        for row_index, row in enumerate(tableArr):
            for col_index, values in enumerate(row):
                for valIndex,value in enumerate(values):
                    if value is not None and target.lower() in value.lower():
                        return [[row_index, col_index,valIndex],value]
        return None

def convertBOLsToDic(bolList):
    bolDic=defaultdict(list)
    for bol in bolList:
        for subArray in bol[3]:
            po_num=subArray[0].split("-")[1].split("-")[0]
            bolDic[po_num].append(subArray)
    return bolDic

def convertDocToPDF(docPath):
    outputPath=docPath.split(".")[0]+".pdf"
    docx2pdf.convert(docPath,outputPath)
    return outputPath