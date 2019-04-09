# -*- coding: utf-8 -*-

from public_en.Excel import OperationExcel

class readExcel(OperationExcel):

    def read_dict_data(self):
        if self.colNum <= 1:
            print("数据行数为1")
        else:
            r = []
            j = 1
            for i in range(self.colNum-1):
                s = {}
                # 从第二行对应values值
                s['rowNum'] = i+2
                values = self.table.row_values(j)
                for x in range(self.colNum):
                    s[self.keys[x]] = values[x]
                r.append(s)
                j += 1
            return r

if __name__ == '__main__':
    aa = "D:\project_address_1\interface4\The_test_case\interfaceTest.xlsx"
    cc = readExcel(aa)
    print(cc)
    print(cc.read_dict_data())
