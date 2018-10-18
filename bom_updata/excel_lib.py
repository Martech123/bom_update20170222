#encoding=utf-8
#author: walker
#date: 2015-09-25
#function: 打印指定access文件、指定表的所有字段


import xlrd
import xlwt
from xlutils.copy import copy



excel_file = r'D:\workspace\Python\bom_updata\123.xls'


def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(r'D:\workspace\Python\bom_edit\cnu_c8811_v1.xlsx')
    # 获取所有sheet
    # 获取所有sheet
    print (workbook.sheet_names()) # [u'sheet1', u'sheet2']


    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0)

    # sheet的名称，行数，列数
    print (sheet.name)
    print (sheet.name,sheet.nrows,sheet.ncols)
    nrows = sheet.nrows
    ncols = sheet.ncols
    print(range(nrows))
    for i in range(nrows):
        print (sheet.row_values(i))
        print ()



    print()
    print()
    print()

def get_coord(n,excel_file):  #坐标索引 返回字符串在EXCEL中的的坐标位置
   #        n---> 'string'
   #        excel_file----> r'D:\*\*.xlsx'
   #
   # workbook = xlrd.open_workbook(r'D:\workspace\Python\bom_edit\cnu_c8811_v1.xlsx')
    workbook = xlrd.open_workbook(excel_file)
    # 获取所有sheet
    # 获取所有sheet
    number_row=0
    number_col=0

    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    #索引“线路板”cell 坐标点   
    for i in range(nrows):
        for j in range(ncols):
            if sheet.cell_value(i,j) ==n:
                number_row=i
                number_col=j
                break

    return number_row,number_col        


def write_excel(lib_q,coord,excel_file):
    w = xlrd.open_workbook(excel_file)
    # a = w.sheet_names()
    # print(a)
    sh = w.sheet_by_index(0)
    wb = copy(w)
    #通过get_sheet()获取的sheet有write()方法
    ws = wb.get_sheet(0)
    ws.write(coord[0],coord[1], lib_q)
    
    wb.save(excel_file)

def setCellCol(row,col,colour,cell_value,excel_file):#cell_value --->str

    # workbook = xlrd.open_workbook(excel_file)
    workbook = xlrd.open_workbook(excel_file,on_demand=True,formatting_info=True)

    wb = copy(workbook)
    ws = wb.get_sheet(0)


    pattern = xlwt.Pattern() # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = colour # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    style = xlwt.XFStyle() # Create the Pattern
    style.pattern = pattern # Add Pattern to Style
    ws.write(row, col, cell_value,style)
    # print('before save')
    wb.save('newbom.xls')
    # print('abc.xls ')
    # workbook = xlwt.Workbook()
    # worksheet = workbook.add_sheet('My Sheet')
    # pattern = xlwt.Pattern() # Create the Pattern
    # pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    # pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    # style = xlwt.XFStyle() # Create the Pattern
    # style.pattern = pattern # Add Pattern to Style
    # worksheet.write(0, 0, 'Cell Contents', style)
    # workbook.save('Excel_Workbook.xls')

if __name__ == '__main__':
    setCellCol(3,3,5,'ssss',excel_file)
    for i in range(9):
        setCellCol(i,3,6,'ssss',excel_file)

