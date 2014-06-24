'''
Created on 2014/6/24

@author: piorpua
'''

class ExcelSheet(object):
    
    def parseSheet(self, sheetName, sheet):
        if (None == sheetName or None == sheet):
            return False
        
        self.__name = sheetName
        
        if (sheet.nrows <= 0 or sheet.ncols <= 0):
            return False
        
        for i in range(0, sheet.ncols):
            self.__dataSize.append(0)
        
        self.__header = sheet.row_values(0)
        for i in range(1, sheet.nrows):
            row = sheet.row_values(i)
            self.__data.append(row)
            
            for j in range(0, sheet.ncols):
                colSize = len(row[j])
                if (colSize > self.__dataSize[j]):
                    self.__dataSize[j] = colSize
        
        return True
    
    def getHeader(self):
        return self.__header
    
    def getDataSize(self):
        return self.__dataSize
    
    def getData(self):
        return self.__data
    
    def getName(self):
        if (None != self.__name):
            return self.__name
        
        return ''
    
    def getRowCount(self):
        if (None != self.__data):
            return self.__data.__len__()
        
        return 0
    
    def getColCount(self):
        if (None != self.__header):
            return self.__header.__len__()
        
        return 0
    
    def isValid(self):
        return self.getName().__len__() != 0 and self.getRowCount() != 0 and self.getColCount() != 0
    
    def show(self):
        print("--- Data ---")
        for i in range(self.getRowCount()):
            print(i, ' : ', self.__data[i])
        
        print()
        print("--- Sheet Name ---")
        print(self.__name, '(', self.getRowCount(), ', ', self.getColCount(), ')')
        print("--- Sheet Header ---")
        print(self.__header)
        print("--- Data Size ---")
        print(self.__dataSize)
    
    def __init__(self):
        self.__name = ''                          # sheet name
        self.__header = []                        # header name (row:0)
        self.__data = []                          # data
        self.__dataSize = []                      # length of data(assist to build sqlite)