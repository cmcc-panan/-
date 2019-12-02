# -*- coding: utf-8 -*-
"""
Created on Mon Aug 26 09:14:22 2019

@author: nocool
"""

import xlrd
import requests

data = xlrd.open_workbook("微信支付申请营业厅汇总0823.xlsx")

table = data.sheet_by_name("Sheet2")

row = table.nrows

for i in range(1,row):
    rowdate = table.row_values(i)
    name = rowdate[1]
    pic_address = rowdate[3]
    pic = requests.get(pic_address,timeout = 15)
    pic_name = name + '.png'
    try :
        with open(pic_name,'wb') as f :
            print("正在下载 : " + pic_name + "的图片!")
            f.write(pic.content)
    except Exception as e:
        print("下载" + name + "时失败了!")
        continue