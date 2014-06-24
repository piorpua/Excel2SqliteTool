'''
Created on 2014/6/24

@author: piorpua
'''

import sqlite3


class DescDB(object):
    
    def create(self, excelData):
        if (None == excelData):
            return False
        
        if (excelData.getSheetCount() <= 0):
            return False
        
        db = sqlite3.connect(self.__path)
        cursor = db.cursor()
        
        try:
            
            sheets = excelData.getSheets()
            for sheetName in sheets.keys():
                sheet = sheets[sheetName]
                if (None == sheet):
                    continue
                
                sql = "drop table if exists " + sheetName + ";";
                cursor.execute(sql)
                
                colCount = sheet.getColCount()
                index = 0
                
                sql = "create table " + sheetName + "(";
                for (colName, dataSize) in zip(sheet.getHeader(), sheet.getDataSize()):
                    index += 1
                    
                    if (index == 1):
                        sql += colName + " text primary key"
                    else:
                        sql += colName + " varchar(" + str(dataSize) + ")"
                        
                    if (index != colCount):
                        sql += ", "
                    else:
                        sql += ");"
                    
                cursor.execute(sql)
                
                for row in sheet.getData():
                    sql = "insert into " + sheetName + " values(";
                    
                    index = 0
                    for value in row:
                        index += 1
                        
                        sql += "\'" + value + "\'"
                        if (index != colCount):
                            sql += ", "
                        else:
                            sql += ");"
                    
                    try:
                        cursor.execute(sql)
                    except:
                        print("Except: ", sql)
            
            db.commit()
            
        finally:
            if (None != cursor):
                cursor.close()
            
        return True
    
    def getPath(self):
        return self.__path

    def __init__(self, dbpath):
        self.__path = dbpath