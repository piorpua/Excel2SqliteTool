'''
Created on 2014/6/24

A tool translate excel(.xls) to sqlite(.db)
Rules:
Sheet Name ---> Table Name
Cell Name ---> Tuple Name

@notice: rely xlrd-0.9.3
@author: piorpua
'''

import xlrd
import os

from data.exceldata import ExcelData
from db.descdb import DescDB

# dir path
dirpath = 'D:\Python\workspace\Excel2SqliteTool'
# excel path
fname = dirpath + '\sample.xls'

# sqlite path
dbpath = dirpath + '\sample.db'

workbook = xlrd.open_workbook(fname)

if __name__ == '__main__':
    excelData = ExcelData()
    if (excelData.parseExcel(workbook)):
        excelData.show()
        
        try:
            os.remove(dbpath)
        except:
            print("\nExcept: remove", dbpath)
            
        descDB = DescDB(dbpath)
        if (descDB.create(excelData)):
            print('\nDB:', dbpath, 'create success.')
        else:
            print('\nDB:', dbpath, 'create failure.')
    else:
        print('\nExcel:', fname, 'parse failure.')
    
    
        