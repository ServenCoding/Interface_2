# -*- coding: utf-8 -*-
# @Time    : 2019/4/9  14:12
# @Author  : Mr lu
# @FileName: Request_.py
# @Software: PyCharm

import requests
from common.Encapsulation_Excel import readExcel
from common.write_Excel import Unit,copy_Excel
import json

def send_request(seeions_data,test_data):
    """
    封装request方法
    :param s:
    :param request:
    :return:
    """
    method = test_data['method']
    url = test_data['url']
    name = test_data['name']
    #请求类型
    type = test_data['type']
    #请求id
    test_nub = test_data['id']
    #判断参数
    #请求url后面的参数patams
    try:
        params = eval(test_data['params'])
    except:
        params = None

    try:
        headers = eval(test_data['headers'])
    except:
        headers = None

    print('************正在执行测试用例:----------%s----------************' % test_nub)
    print('用例名称:%s' % name)
    print('请求方式:%s,请求Url:%s' % (method,url))
    print('请求参数:%s' % params)

    #post请求body内容
    try:
        bodydata = exec(test_data['body'])
    except:
        bodydata = {}

    #判断传入Data数据类型
    if type == 'data':
        body = bodydata
    elif type == 'json':
        body  = json.dumps(bodydata)
    else:
        body = bodydata

    #判断传入的数据类型
    if method == 'post':
        print("Post请求类型为：%s,Body内容为:%s"% (type,bodydata))
    elif method == 'get':
        print("Get请求类型为：%s,Body内容为:%s"% (type,bodydata))
    elif method == 'put':
        print("Get请求类型为：%s,Body内容为:%s" % (type, bodydata))
    elif method == 'delete':
        print("Get请求类型为：%s,Body内容为:%s" % (type, bodydata))

    verify = False #避免ssl认证
    res = {}

    try:
        r = seeions_data.request(method=method,data='',url=url,headers='',verify=verify,params='')
        print("页面返回信息:%s" % r.content.decode('utf-8'))
        res['id'] = test_data['id']
        res['rowNum'] = test_data['rowNum']
        # res['result'] = test_data['result']
        res['statuscode'] = str(r.status_code)
        res['text'] = r.content.decode('utf-8')
        res['times'] = str(r.elapsed.total_seconds())
        if  res['statuscode'] == "200":
            res['error'] = ""
            res['msg'] = ""
            if test_data['checkpoint'] in res['text']:
                res['result'] = "pass"
                print("用例结果:%s------------->%s" % (test_nub,res['result']))
            else:
                res['result'] = "fail"
                print("用例结果:%s------------->%s" % (test_nub,res['result']))
            return res
        else:
            res['error'] = res['text']
    except Exception as msg:
        res['msg'] = str(msg)
        return res

def write_result(result,filename):
    #返回结果的行数
    row_hub = result['rowNum']
    #写入Excel中
    wt = Unit(filename)
    wt.write_Sheets(row_hub,9,result['statuscode'])
    wt.write_Sheets(row_hub,10, result['times'])  # 耗时
    wt.write_Sheets(row_hub,11, result['error'])  # 状态码非200时的返回信息
    wt.write_Sheets(row_hub,12, result['result'])  # 测试结果 pass 还是fail
    wt.write_Sheets(row_hub,13, result['msg'])  # 抛异常



if __name__ == "__main__":
    data = readExcel('D:\project_address_1\interface4\The_test_case\interfaceTest.xlsx').read_dict_data()
    print(data[0])
    s = requests.session()
    res = send_request(s,data[0])
    copy_Excel("D:\project_address_1\interface4\The_test_case\interfaceTest.xlsx","D:\project_address_1\interface4\The_test_case\q2.xlsx")
    write_result(res, filename="D:\project_address_1\interface4\The_test_case\q2.xlsx")
