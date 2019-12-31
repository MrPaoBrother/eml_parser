# -*- coding:utf8 -*-

import email
import re
import os

import xlwt
from lxml import etree

"""
    相关配置
"""

root_path = "./data/"
save_file= "./result/result.xls"

sheet_name = "sheet1"

# 某个关键词后面的所有列
key_word = u"有效响应率"

def fetch_filenames():
    result = []
    files = os.listdir(root_path)
    for file in files:
        if u".eml" in file.decode("utf8"):
            result.append(file.decode("utf8"))
    return result

def parser(filename):
    fp = open(root_path + filename, 'r')
    message = email.message_from_file(fp)
    for item in message.walk():
        html = item.get_payload(decode=True)
        dom = etree.HTML(html, parser=etree.HTMLParser(encoding="utf-8"))
        product = ""
        start_record = False
        client_name = filename.split(' ')[0]

        trs = dom.xpath('//table')[0].xpath('//tr')
        heads = trs[0].xpath('.//td')

        for head in heads:
            txt = head.xpath('string(.)')
            if key_word in txt:
                start_record = True
                continue
            if start_record:
                product += txt + "\n"
            
        search_count = trs[1].xpath('.//td')[1].xpath('string(.)')
        
        return [client_name, int(search_count), product]

def write_xls(data_list):
    # 每次生成就重写一次 保证数据正确性
    excel_w = xlwt.Workbook(encoding='utf-8')
    excel_w_sheet = excel_w.add_sheet(sheet_name)

    head = data_list[0]
    # 写head
    for i in range(len(head)):
        excel_w_sheet.write(0, i, label=head[i])

    # 写body
    for i in range(1, len(data_list)):
        for j in range(len(head)):
            excel_w_sheet.write(i, j, label=data_list[i][j])

    excel_w.save(save_file)

def init(need_regenerate = True):
    if need_regenerate:
        if os.path.exists(save_file):
            os.remove(save_file)

def process():
    init()
    data_list = []
    head = ['客户姓名', '调用总量(两个月)', '产品']
    data_list.append(head)
    process = 0
    eml_files = fetch_filenames()
    for eml_file in eml_files:
        if process % 5 == 0:
            print ("process: %s | %s" % (process, len(eml_files)))
        result = parser(eml_file)
        data_list.append(result)
        process += 1
    write_xls(data_list)

if __name__ == '__main__':
    try:
        process()
    except Exception as e:
        print "error: %s" % str(e)