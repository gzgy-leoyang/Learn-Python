import os
import sys
import xlrd
import xlwt


# 定义输出文件描述
class Output_file:
    file_name   ='book.xls' # 输出文件名称
    sheet_name  ='sheet1'   # 文件中的表格名称
    title       =[]         # 表格的标题行
    col_index   =[]         # 需要插入的列的编号组
    col_list    =[]         # 需要插入的列的内容组

    # 向输出文件添加一列内容，包括列的编号和值内容
    def add_col(self,_col_num,_col_val):
        self.col_list.append( _col_val )
        self.col_index.append( _col_num )
# 输出文件定义完成


# 抓取指定列，该列来自指定文件，指定表格和列编号
def get_col_by_book_sheet_num( _file_name ,_sheet_number,_col_number,_start_row=1,_end_row=None ):
    workbook=xlrd.open_workbook( _file_name )
    worksheet=workbook.sheet_by_index( _sheet_number )
    return worksheet.col_values( _col_number ,_start_row,_end_row )

# 将拉取的列数据写入新文件，新表格的指定列位置，
def create_new_file( _outfile ):
    new_book    = xlwt.Workbook( encoding ="utf-8",style_compression = 0)
    sheet_0     = new_book.add_sheet( _outfile.sheet_name ,cell_overwrite_ok=True)
    # 添加标题行
    for i in range(0,len( _outfile.title)):
        sheet_0.write(0 ,i ,_outfile.title[i] )

    # 将拉出来的列写入目标文件
    for j in range(0,len( _outfile.col_list)):
        for i in range(0,len( _outfile.col_list[j] )):
            sheet_0.write(i+1,_outfile.col_index[j],_outfile.col_list[j][i] )
    # 写入文件
    new_book.save( _outfile.file_name ) 
########################################33

# 主函数
if __name__ == "__main__":
    out_file            = Output_file()
    out_file.file_name  ='BBook.xls'
    out_file.sheet_name ='sheet1'
    out_file.title      =['姓名','年龄','学历']

    # 提取 'Book1.xls'文件的 第0表格 第0列 的内容，添加到输出文件的第0列
    col_name = get_col_by_book_sheet_num( 'Book1.xls',0,0 )
    out_file.add_col(0,col_name )

    # 提取 'Book1.xls'文件的 第0表格 第1列 的内容，添加到输出文件的第1列
    col_age = get_col_by_book_sheet_num( 'Book1.xls',0,1 )
    out_file.add_col(1,col_age )

    # 提取 'Book2.xls'文件的 第0表格 第1列 的内容，添加到输出文件的第2列
    col_degree = get_col_by_book_sheet_num( 'Book2.xls',0,1 )
    out_file.add_col(2,col_degree )
    
    # 按照已经添加的列组，生成文件
    create_new_file( out_file )