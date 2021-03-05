import xlrd
from xlrd import xldate_as_tuple
import datetime
class ExcelData():
    # 初始化方法
    def __init__(self, data_path,baseData):
        #定义一个属性接收文件路径
        self.data_path = data_path
        # 定义一个属性接收工作表名称
        # self.sheetname = sheetname
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheets()[0]
        # 根据工作表的索引获取工作表的内容（方式②）
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(0)
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        # 获取工作表的有效列数
        self.colNum = self.table.ncols
        self.baseData = baseData

    # 定义一个读取excel表的方法
    def readExcel(self):
        # 定义一个空列表
        datas = {}
        for i in range(1, self.rowNum):
            c_cell = self.table.cell_value(i, 4)
            datas[c_cell]=c_cell
            # print(c_cell)
        return datas

        # 计算不在data_path的中的数据。
    def readExcelAndcalc(self):
        # 定义一个空列表
        datas = []
        f=self.data_path +".txt"
        for i in range(1, self.rowNum):
            c_cell = self.table.cell_value(i, 4)
            # d[c_cell]=c_cell
            if self.baseData.get(c_cell,-1) != -1:
                datas.append(c_cell)
                with open(f,"a") as file:  
                    file.write(c_cell+"\n")
        return datas
if __name__ == "__main__":
    data_path = "./20210120-xxxx.xlsx"
    get_data = ExcelData(data_path,"")
    datas = get_data.readExcel()
    dataschecks =["北峰.xlsx","大街.xlsx","旧治.xlsx","沙圪塔.xlsx","营镇.xlsx","金滩镇.xlsx","龙王庙.xlsx","埝头.xlsx","孙甘店.xlsx","束馆镇.xlsx","王村乡.xlsx","西付集.xlsx","铺上乡.xlsx", "万堤.xlsx","大名镇.xlsx","张集 .xlsx","杨桥.xlsx","红庙乡.xlsx","西未庄.xlsx","黄金堤.xlsx"]
    # dataschecks =["北峰.xlsx"]
    lendatas = len(dataschecks)
    print(lendatas)
    for datascheck in dataschecks:
        cur_data = ExcelData(datascheck,datas)
        #查询并写入记事本
        cur_datas = cur_data.readExcelAndcalc()

