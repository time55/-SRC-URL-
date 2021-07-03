#!/usr/bin/python3
# -*- coding: utf-8 -*-

# coding: utf-8
# Team : Network security group 
# Author：ych0515
# Date ：2021/7/2 23:56
# Tool ：PyCharm

import requests
import re
import time
import os
import xlrd
import warnings
import ssl
import traceback
import xlwt
import time
from xlutils.copy import copy
from bs4 import BeautifulSoup
warnings.filterwarnings("ignore")
ssl._create_default_https_context = ssl._create_unverified_context


headers = {

    'Host':'www.butian.net',
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
    'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language':'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
    'Accept-Encoding':'gzip, deflate',
    'Connection':'close',
    'Referer':'https://www.butian.net/Reward/plan/2',
    'Cookie':'登录后的cookie值',
    'Upgrade-Insecure-Requests':'1',
    'Pragma':'no-cache',
    'Cache-Control':'no-cache'
    }

def parse_data(jsons):
    datas = (jsons['data'])
    real_data = (datas['list'])
    try:
        for d in real_data:
            # 获取厂商url
            url = "https://www.butian.net/Loo/submit?cid="+d['company_id']
            webdata = requests.get(url=url,headers=headers)
            webdata.encoding = webdata.apparent_encoding
            soup = BeautifulSoup(webdata.text, 'html.parser')
            input_value = soup.find_all('input',class_='input-xlarge')[1]['value']
            value_final = []
            value_final_list = []
            value_final.append(d['company_name'])

            value_final.append(input_value)
            value_final_list.append(value_final)
            print(d['company_name'] + "    " + input_value)
            create_excel(value_final_list,0)
    except Exception as e:
        traceback.print_exc()



# 保存数据库
def create_excel(value,orderNum):
    # 创建excel表
    excel_path = "补天公益SRC厂商URL.xls"
    if os.path.exists(excel_path):
        print("表格已创建")
    else:
        # 创建excel表
        workbook = xlwt.Workbook()
        sheet1 = workbook.add_sheet("公益SRC名称URL")
        sheet1.write(0, 0, label='厂商名称')
        sheet1.write(0, 1, label='厂商URL')
        workbook.save(excel_path)
    # 保存至excel表
    write_excel_xls_append(excel_path, value,orderNum)


# excel表中追加数据
def write_excel_xls_append(path, value,orderNum):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[orderNum])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(orderNum)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")

# 截止2021-07-03，补天公益src厂商为186页，如有变化，可修改以下循环中
def main():
    for i in range(1,187):
        print('目前获取第 {} 页'.format(i))
        url = 'https://www.butian.net/Reward/pub'
        data = {
            's': '1',
            'p': i
        }
        r = requests.post(url=url, data=data)
        try:
            parse_data(r.json())
        except Exception as e:
            traceback.print_exc()
            print(e)

if __name__ == '__main__':
    main()
    print("补天公益SRC厂商URL获取结束...")