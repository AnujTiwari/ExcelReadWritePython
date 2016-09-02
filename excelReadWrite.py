import xlrd
import xlwt

from tempfile import TemporaryFile
from xlwt import Workbook
#----------------------------------------------------------------------

def open_file(path):
    """
    Open and read and write an Excel file
    """
    book = xlrd.open_workbook(path)

    sheet = book.sheet_by_index(0)

    print sheet.name
    print sheet.ncols
    print sheet.nrows

    dataList = []

    for row_index in range(sheet.nrows):
         for col_index in range(sheet.ncols):
     
             #print col_index
             #print row_index
             print sheet.cell(row_index,col_index).value
             if (col_index == 0):
                 value1 = sheet.cell(row_index,col_index).value
                 dataList.append(value1)
             if (col_index == 1):
                 value2 = sheet.cell(row_index,col_index).value
                 dataList.append(value2)
                 print "value1*value2 : " + str((value1*value2)/36000)
                 dataList.append(str((value1*value2)/36000))
                 
    #----------------------------------------------------------------------            

    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')

    for index, elem in enumerate(dataList):
        print(index, elem)

        if  (index%3 == 1):
             sheet1.write(index/3,1,elem)
        if  (index%3 == 2):
             sheet1.write(index/3,2,elem)
        if  (index%3 == 0):
             sheet1.write(index/3,0,elem)


    sheet1.col(0).width = 2000

    book.save('simple.xls')
    book.save(TemporaryFile())

   
 
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "J:\\Book2.xls"
    open_file(path)
