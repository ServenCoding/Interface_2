# -*- coding: utf-8 -*-
# @Time    : 2019/1/3  20:16
# @Author  : MrLu！！
# @FileName: Test_case.py
# @Software: PyCharm

import os
import unittest

import ddt
import requests

requests.packages.urllib3.disable_warnings()
from common.Log import Logger
from common.Encapsulation_Excel import readExcel
from common.write_Excel import copy_Excel
from common.Request_ import send_request

#日志
logger =Logger(logger='testCase').getlog()
#获取old_Excel路径
old_Ex = os.path.dirname(os.getcwd()) + '\The_test_case\interfaceTest.xlsx'
#获取new_Excel路径
new_Ex = os.path.dirname(os.getcwd()) + '\The_test_case\q2.xlsx'
#执行Excel文件
testdata = readExcel(old_Ex).read_dict_data()

@ddt.ddt
class jie(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.s = requests.session()
        logger.info("开始测试")
        copy_Excel(old_Ex,new_Ex)
        logger.info("新旧表拷贝完成")

    @ddt.data(*testdata)
    def test_Api(self,data):
        '''
        接口测试Case
        :return:
        '''
        requests.packages.urllib3.disable_warnings()
        res = send_request(self.s,data)
        check = data['checkpoint']
        print("检查点:->%s" %check)
        res_text = res['text']
        print("返回的实际结果:->%s" %res_text)

    @classmethod
    def tearDownClass(cls):
        logger.info("测试结束")

if __name__ == "__name__":
    unittest.main()