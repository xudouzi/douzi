#encoding=utf-8
#影院信息的更新
import unittest
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
from excel import *
import time
import csv
import codecs
import xlrd
import string
import xlsxwriter
import xlwt

datapath_reference = r"/Users/douzi/Desktop/ge/副本5013724-MASSIVE_MIMO_FMC——焊接bom (1).xlsx"
sheetname_reference = r"MASSIVE_MIMO_FMC"

datapath_nonreference = r"/Users/douzi/Desktop/ge/MASSIVE_MIMO_FMC——焊接bom20190618.xlsx"
sheetname_nonreference ="MASSIVE_MIMO_FMC"

def sheet_name(dataPath,sheetname):
    parseExcel = PsrseExcel()
    parseExcel.loasworkBook(dataPath)
    sheet = parseExcel.getSheetByName(sheetname)
    return sheet,parseExcel


#excel 表中读取位号
def read_weihao_excel(rowNo,colNo,datapath,sheetname):
    item=sheet_name(datapath,sheetname)
    sheet = item[0]
    parseExcel = item[1]
    data=parseExcel.getCellValue(sheet,rowNo=rowNo,colNo=colNo)
    return data
data_non = []
data_ref = []
non_excel = read_weihao_excel(13,2,datapath_nonreference,sheetname_nonreference).replace('，', ',').split(',')
ref_excel = read_weihao_excel(13,2,datapath_reference,sheetname_reference).replace('，', ',').split(',')
for i in range(len(non_excel)):
     data_non.append(str(non_excel[i]))

for i in range(len(ref_excel)):
     data_ref.append(str(ref_excel[i]))

item = list(set(non_excel).difference(set(ref_excel)))
print data_non,data_ref,item




col=2
def max_row(datapath,sheetname):
    item = sheet_name(datapath,sheetname)
    sheet = item[0]
    parseExcel = item[1]
    max_row = parseExcel.getRowEndVumber(sheet)
    return max_row
ref_max_row = max_row(datapath_reference,sheetname_reference)
nonoref_max_row = max_row(datapath_nonreference,sheetname_nonreference)

max_row_excel = nonoref_max_row

for i in range(1, max_row_excel):
     reference_excel = read_weihao_excel(i + 1, col, datapath_reference, sheetname_reference).replace('，', ',')
     nonreference_excel = read_weihao_excel(i + 1, col, datapath_nonreference, sheetname_nonreference).replace('，', ',')

     item = list(set(nonreference_excel).difference(set(reference_excel)))
     # print item
     if item != []:
          print i + 1, item
