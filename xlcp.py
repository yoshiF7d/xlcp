import argparse
import xlwings 
import os
import shutil
import numpy as np
from enum import Enum,auto

class style():
        BLACK = '\033[30m'
        RED = '\033[31m'
        GREEN = '\033[32m'
        YELLOW = '\033[33m'
        BLUE = '\033[34m'
        MAGENTA = '\033[35m'
        CYAN = '\033[36m'
        WHITE = '\033[37m'
        UNDERLINE = '\033[4m'
        RESET = '\033[0m'

class Excel():
    def __init__(self):
        self.app = xlwings.App(visible=False)
        self.books = []
        for app in xlwings.apps:
            if app == self.app:
                continue
            for book in app.books:
               self.books.append(book)

    def __del__(self):
        self.app.quit()

    def isOpen(self,file):
        if file in [book.fullname for book in self.books]:
            return True
        return False

    def open(self,file):
        for book in self.books:
            if book.fullname == file:
                return book
        book = self.app.books.open(file)
        self.books.append(book)
        return book

    def close(self,book):
        if book not in self.books:
            book.close()

class Position(Enum):
    TOP = auto()
    BOTTOM = auto()
    RIGHT = auto()
    LEFT = auto()
    
    def get(string):
        for item in Position:
            if item.name == string.upper():
                return item
        return None

    @staticmethod
    def parse(string):
       s = string.strip().strip('{[()]}').split(',')
       s1 = s[0]
       s2 = s[1] if len(s) > 1 else None
       return list(map(Position.get,(s1,s2)))

class Range:
    def __init__(self,string):
        self.startRow = 0
        self.startColumn = 0
        self.endRow = 0
        self.endColumn = 0
        self.range = None
        self.parse(string)

    def __str__(self):
        return '{{{0},{1},{2},{3}}}'.format(self.startRow,self.startColumn,self.endRow,self.endColumn)
 
    def set(self,sheet):
        tl = sheet[self.startRow,self.startColumn]
        if not self.endRow:
            bl = tl.end('down')
        else:
            bl = tl.offset(self.endRow - self.startRow,0)

        if not self.endColumn:
            br = bl.end('right')
        else:
            br = bl.offset(0,self.endColumn - self.startColumn)
        
        self.range = sheet.range(tl,br)

    def getHeaderCell(self,position):
        if self.range:
            (row,column) = self.range.shape
            tl = self.range[0]
            
            ofsr = 0
            ofsc = 0
            
            if position[0] == Position.TOP:
                ofsr = -1
            elif position[0] == Position.LEFT:
                ofsc = -1
            elif position[0] == Position.BOTTOM:
                ofsr = row
            elif position[0] == Position.RIGHT:
                ofsc = column

            if position[1]:
                if position[1] == Position.RIGHT:
                    ofsc = column - 1
                elif position[1] == Position.BOTTOM:
                    ofsr = row - 1 

            return tl.offset(ofsr,ofsc)

    def parse(self,string):
        if type(string) is str:
            (start,end) = string.split(":")
            (self.startRow,self.startColumn) = Range.parseCell(start)
            (self.endRow,self.endColumn) = Range.parseCell(end)
   
    @staticmethod
    def parseAlpha(astr):
        num=0
        exp=1
        if not astr:
            num = None
        else:
            for a in astr:
                num = exp*num + ord(a.lower())-96
                exp *= 26
        if num:
            num = num - 1
        return num

    @staticmethod
    def parseCell(cstr):
        (alpha,num) = ("","")
        for i in range(len(cstr)):
            if cstr[i].isalpha():
                alpha += cstr[i]
            else:
                num = cstr[i:]
                break
        else:
            num = ""
        
        if num :
            num = int(num) - 1
        else:
            num = None
        return (num,Range.parseAlpha(alpha))

class SubOrder:
    def __init__(self,order,args):
        (
            self.header,
            self.filein,
            self.sheetin,
            self.rangein,
            self.fileout, 
            self.sheetout, 
            self.rangeout
        ) = args
        
        self.rangein = Range(self.rangein)
        self.rangeout = Range(self.rangeout)
        self.order = order

    def __str__(self):
        return ('header : {0}\n'
                'filein : {1}\n'
                'sheetin : {2}\n'
                'rangein : {3}\n'   
                'fileout : {4}\n' 
                'sheetout : {5}\n'
                'rangeout : {6}\n'
                ).format(self.header,self.filein,self.sheetin,str(self.rangein),self.fileout,self.sheetout,str(self.rangeout)) 
    def isProper(self):
        return all([self.filein,self.sheetin,self.fileout,self.sheetout])

    def read(self):
        filein = os.path.join(self.order.dirin,self.filein)
        if not os.path.exists(filein):
            print(style.CYAN + self.filein + style.RESET + ' does not exist in ' + style.CYAN + self.order.dirin + style.RESET)
            return False
        
        print('loading ' + style.CYAN + filein + style.RESET)
        bookin = order.excel.open(filein)
        if self.sheetin not in [sheet.name for sheet in bookin.sheets]:
            print(style.CYAN + self.sheetin + style.RESET + ' does not exist in ' + style.CYAN + filein + style.RESET)
            order.excel.close(bookin)
            return False
        
        sheetin = bookin.sheets[self.sheetin]
        self.rangein.set(sheetin)
        self.order.array = self.rangein.range.options(convert=np.array,ndim=2).value
        if self.order.transpose:
            self.order.array = self.order.array.transpose()
        print('data shape : ' + style.CYAN + str(self.order.array.shape) + style.RESET)
        order.excel.close(bookin)
        return True

    def write(self):
        fileout = os.path.join(self.order.dirout,self.fileout)
        if fileout in self.order.dict:
            bookout = self.order.dict[fileout]
        else:
            if not os.path.exists(fileout):
                shutil.copy(self.order.temp,fileout)
            elif not args.forceOverwrite:
                strin=input('file ' + style.CYAN +  f'{fileout}' + style.RESET + 
                        ' already exists. Overwrite? (' + style.CYAN + 'y' + style.RESET + '/' + style.RED + 'n' + style.RESET + ')')
                if strin[0] != 'y': 
                    order.nolist.append(self.fileout)
                    return

            bookout = order.excel.open(fileout)
        
        self.order.dict[fileout] = bookout
        if self.sheetout not in [sheet.name for sheet in bookout.sheets]:
            bookout.sheets.add(self.sheetout)
            print('created' + style.CYAN + self.sheetout + style.RESET + ' in ' + style.CYAN + self.fileout)

        sheetout = bookout.sheets[self.sheetout] 
        self.rangeout.set(sheetout)
        
        if self.header:
            self.rangeout.getHeaderCell(self.order.headerPosition).value = self.header

        rangeout = self.rangeout.range
        rangeout.value = self.order.array[:rangeout.shape[0],:rangeout.shape[1]] 

class Order:
    def __init__(self,sheet,args,excel):
        self.temp = None 
        self.dirin = None 
        self.dirout = None 
        self.list = []
        self.nolist = []
        self.dict = {}

        self.headerPosition = Position.parse(args.headerPosition)
        self.transpose = args.transpose
        self.load(sheet)
        self.excel = excel
    
    def __str__(self):
        string = (
        'temp : {0}\n'
        'dirin : {1}\n'
        'dirout : {2}\n\n'
        ).format(self.temp,self.dirin,self.dirout)
        for elem in self.list:
            string += str(elem)
            if elem is not self.list[-1]:
                string += '\n\n'
        return string

    def load(self,sheet):
        self.temp = sheet.range('B1').value
        self.dirin = sheet.range('B2').value
        self.dirout = sheet.range('B3').value
        self.list = []
        self.dict = {}
        self.nolist = []
        tl = sheet.range('A6')
        bl = tl.end('down')
        br = bl.offset(0,6)
        
        data = sheet.range(tl,br).value
        for row in data:
            subOrder = SubOrder(self,row)
            if subOrder.isProper():
                self.list.append(subOrder)
                print(subOrder)

    def execAll(self):
        for suborder in self.list: 
            if suborder.fileout in self.nolist:
                continue
            print(style.CYAN + suborder.filein + style.RESET + '->' + style.CYAN + suborder.fileout + style.RESET)
            
            if suborder.read():
                suborder.write()
        for bookname,book in self.dict.items():
            print('saving ' + style.CYAN + bookname + style.RESET)
            book.save()
            print('saved')
            self.excel.close(book)

parser = argparse.ArgumentParser()
parser.add_argument('orderFile',help='input order file')
parser.add_argument('-t','--transpose',action='store_true',help='transpose data')
parser.add_argument('-f','--forceOverwrite',action='store_true',help='force overwrite')
parser.add_argument('--headerPosition',default='(top,left)',help='position of the header with respect to the data. default : (top,left). set this to (left,top) to put header to the side.')
args = parser.parse_args()

os.system('')

try:
    excel = Excel()
    book = excel.open(args.orderFile)
    orderList = []
    for sheet in book.sheets:
        orderList.append(Order(sheet,args,excel))
    excel.close(book)

    for order in orderList:
        order.execAll()

finally:
   pass 
