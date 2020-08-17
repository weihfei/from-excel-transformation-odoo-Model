import xlrd
# from xlrd import xldate_as_tuple
import datetime


class ExcelData():
    def __init__(self):
        self.dataPath = input('请输入文件地址')
        self.data = xlrd.open_workbook(self.dataPath)
        self.table = self.data.sheet_by_name('Sheet1')
        self.rowNum = self.table.nrows
        self.colNum = self.table.ncols
        # self.getRow = input("请输入起始行号")
        # self.getCol = input('请输入起始列号')
        # self.field = ["field.Char", "field.Float", "field.Integer", "field.Datetime"]

    def readExcel(self):
        # 声明保存的空集合
        odooData = []
        for i in range(1, self.rowNum):
            for j in range(1, self.colNum):

                if j == self.colNum - 1:
                    continue
                if j == 1:
                    item_data = self.table.row(i)[j].value + '='
                if j == 2:
                    item_data += "field.Char(" + "'string'='" + self.table.row(i)[j].value + "')\n"

            odooData.append(item_data)

        print(odooData)
        f = open('odooModel.txt', mode='w')
        f.writelines(odooData)
        f.close()


if __name__ == '__main__':
  excelData = ExcelData()
  excelData.readExcel()

