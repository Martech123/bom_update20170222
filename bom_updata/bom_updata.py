#encoding=utf-8
#author: walker
#date: 2015-09-25
#function: 打印指定access文件、指定表的所有字段


import access_lib
import excel_lib
import xlrd

excel_file = r'D:\workspace\Python\bom_edit\cnu_c8811_v2.xls'

access_file = r'D:\workspace\Python\bom_edit\credo_cis_lib.mdb'
# access_file = r'Driver={Microsoft Access Driver (*.mdb)};DBQ=D:\workspace\Python\bom_edit\credo_cis_lib.mdb'


def bom_updata(access_file_in,excel_file_in):
    print('\\\\\\\\\\\\\\\\START Heare\\\\\\\\\\\\\\\\\\\\\\' )
    print()
    print()
    access_file = r'Driver={Microsoft Access Driver (*.mdb)};DBQ='+access_file_in
    excel_file = excel_file_in
    print('******************************')
    print('access_flie=',access_file)


    excel_number_row=0
    excel_number_col=0
    excel_ID_cell = 0,      #ID坐标
    excel_nq_cell = 0,0     #库存数量坐标
    q_access_lib = 0,0

    access_ID=''


    workbook = xlrd.open_workbook(excel_file)
    print(workbook.sheet_names())
    sheet = workbook.sheet_by_index(0)
    excel_number_row = sheet.nrows
    excel_number_col = sheet.ncols

    excel_ID_cell = excel_lib.get_coord('物料编码',excel_file)
    excel_nq_cell = excel_lib.get_coord('库存数量',excel_file)
    print(excel_ID_cell)
    print(excel_nq_cell)



    for i in range(excel_number_row-excel_ID_cell[0]-1):
        if sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])!='':
            access_ID = str(int(sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])))
            q_access_lib = access_lib.read_access_quantity(access_ID,access_file)
            print(q_access_lib)
            if q_access_lib[0]==1:
                excel_lib.write_excel(q_access_lib[1],[i+1+excel_ID_cell[0],excel_nq_cell[1]],excel_file)

            # print(access_ID,access_lib.read_access_quantity(access_ID,access_file))
            # print(access_ID)

        else:
            access_ID=''
            print('当前ID 为NULL')

def DB_updata(access_file_in,excel_file_in):
    print('\\\\\\\\\\\\\\\\DB updata START Heare\\\\\\\\\\\\\\\\\\\\\\' )
    print()
    print()
    access_file = r'Driver={Microsoft Access Driver (*.mdb)};DBQ='+access_file_in
    excel_file = excel_file_in
    
    


    excel_number_row=0
    excel_number_col=0
    excel_ID_cell = 0,      #ID坐标
    excel_nq_cell = 0,0     #库存数量坐标
    q_access_lib = 0,0
    q_excel = 0
    nq_access = 0

    access_ID=''
    excel_cell_value = ''


    workbook = xlrd.open_workbook(excel_file)
    print(workbook.sheet_names())
    sheet = workbook.sheet_by_index(0)
    excel_number_row = sheet.nrows
    excel_number_col = sheet.ncols

    excel_ID_cell = excel_lib.get_coord('物料编码',excel_file)
    excel_nq_cell = excel_lib.get_coord('需求量',excel_file)
    print(excel_ID_cell)
    print(excel_nq_cell)



    for i in range(excel_number_row-excel_ID_cell[0]-1):
        if sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])!='':
            access_ID = str(int(sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])))
            q_excel = int(sheet.cell_value(i+1+excel_nq_cell[0],excel_nq_cell[1]))
            print(q_excel)
            print(type(q_excel))
            q_access_lib = access_lib.read_access_quantity(access_ID,access_file)
            # print(q_access_lib)
            if q_access_lib[0]==1 :
                if str(q_access_lib[1])!= 'Y' and str(q_access_lib[1])!= 'N' and q_access_lib[1]!='':
                    nq_access = int(q_access_lib[1]) - q_excel
                    print ('updta db value = ',nq_access)
                    access_lib.update_access_quantity(access_ID,str(nq_access),access_file)
                else:
                    print('当前ID 在数据库中数量为空 或 Y or N')

                    excel_cell_value =str(int(sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])))

                    excel_lib.setCellCol(i+1+excel_ID_cell[0],excel_ID_cell[1],5,excel_cell_value,excel_file)
            else:
                print('当前ID 数据库不存在 或者 无ID')
                excel_cell_value =str(int(sheet.cell_value(i+1+excel_ID_cell[0],excel_ID_cell[1])))
                excel_lib.setCellCol(i+1+excel_ID_cell[0],excel_ID_cell[1],5,excel_cell_value,excel_file)



        else:
            
            access_ID=''
            print('当前ID 为NULL')
            







if __name__ == '__main__':
    up = DB_updata(access_file,excel_file)























