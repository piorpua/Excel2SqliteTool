'''
Created on 2014/6/24

@author: piorpua
'''
from data.excelsheet import ExcelSheet

class ExcelData(object):
    
    def parseExcel(self, workbook):
        if (None == workbook):
            return False
        
        sheetNames = workbook.sheet_names()
        if (None == sheetNames):
            return False
        
        sheetCount = sheetNames.__len__()
        if (sheetCount == 0):
            return False
        
        for i in range(sheetCount):
            sheet = ExcelSheet()
            sheetName = sheetNames[i]
            if (sheet.parseSheet(sheetName, workbook.sheet_by_index(i))):
                self.__sheets[sheetName] = sheet
        
        return True
    
    def getSheets(self):
        return self.__sheets
    
    def getSheetCount(self):
        return self.__sheets.__len__()
    
    def show(self):
        for key in self.__sheets.keys():
            sheet = self.__sheets[key]
            if (None == sheet):
                continue
            
            print()
            sheet.show()
            print()

    def __init__(self):
        self.__sheets = {}                # 表名:ExcelSheet