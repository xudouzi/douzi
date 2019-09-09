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
from openpyxl.styles import Alignment

datalist=[]




if __name__ == "__main__":
    compareData()