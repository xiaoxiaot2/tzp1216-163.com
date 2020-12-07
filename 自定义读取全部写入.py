import xlrd
import xlwt
import datetime
from xlrd import xldate_as_tuple
'''
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1，所以我们必须在程序中做判断处理转换
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''
actrow = [3,6]
actcol = [1,2,3,4,5,6,7,10] 
rowSet = [m - 1 for m in actrow]
colSet = [n - 1 for n in actcol]     
data_path = "G:\L1216.xlsx"
sheetname = "Sheet1"
class Open_ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
         #定义一个属性接收文件路径
        self.data_path = data_path
        # 定义一个属性接收工作表名称
        self.sheetname = sheetname
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_name(self.sheetname)
        # 根据工作表的索引获取工作表的内容（方式②）
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(1)
        # 获取工作表的有效行数
        self.maxrowNum = self.table.nrows
        # 获取工作表的有效列数
        self.maxcolNum = self.table.ncols
        
    def readExcel(self):
        # 定义一个空列表
        datas = []
        for i in rowSet:
            # 定义一个空字典
            #sheet_data = {}
            for j in colSet:
                # 获取单元格数据类型
                c_type = self.table.cell(i,j).ctype
                # 获取单元格数据
                c_cell = self.table.cell_value(i, j)
                if c_type == 2 and c_cell % 1 == 0:  # 如果是整形
                    c_cell = int(c_cell)
                elif c_type == 3:
                    # 转成datetime对象
                    date = datetime.datetime(*xldate_as_tuple(c_cell,0))
                    c_cell = date.strftime('%Y/%d/%m %H:%M:%S')
                elif c_type == 4:
                    c_cell = True if c_cell == 1 else False
                elif c_type == 0:
                    c_cell = '*'
               # sheet_data[self.keys[j]] = c_cell
                # 循环每一个有效的单元格，将字段与值对应存储到字典中
                # 字典的key就是excel表中每列第一行的字段
                # sheet_data[self.keys[j]] = self.table.row_values(i)[j]
            # 再将字典追加到列表中
                datas.append(c_cell)
        # 返回从excel中获取到的数据：以列表存字典的形式返回
        return datas

if __name__ == "__main__": 
    get_data = Open_ExcelData(data_path, sheetname)
    datas = get_data.readExcel()
    wb = xlwt.Workbook(encoding='utf-8') 
    ws = wb.add_sheet('write_after', cell_overwrite_ok=True) 
    # 写入excel
    # 开始添加数据
     # 参数对应 行, 列, 值
    row, col ,k= 1, 0, 0
    for i in datas:
           # 指定行、列的单元格，添加数据
        if(k<len(colSet)):
            ws.write(row, col, i)
            # 列增加
            col += 1
            k+=1
        if(k == len(colSet)):
            row+=1
            col=0 
            k=0
        #    保存
    wb.save('demo.xls')
    print(datas)
