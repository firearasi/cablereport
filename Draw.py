import os
from datetime import date, datetime, time
from babel.dates import format_date, format_datetime, format_time

import dateutil.parser as dateparser
import argparse
import numpy as np
import matplotlib.pyplot as plt

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.shared import Inches


#还要分端口改
def read_data_from_file(fname):
    with open(fname, 'r') as file:
        lines = file.readlines()
        lines=[s.strip() for s in lines]
        result = []
        data_lines=list(range(12, len(lines),2))
        time_stamp = lines[4][6:]
        time=dateparser.parse(time_stamp)
        for l in data_lines:
            result.append(lines[l].split(None)[1:])
        return (time, np.array(result, dtype='float16'))
        

parser = argparse.ArgumentParser(description='Data path name.')

#命令行参数
parser.add_argument('path', help='数据文件夹')
parser.add_argument('--product', help='产品名称')
parser.add_argument('--samplenumber', help='样品编号')
parser.add_argument('--institution', help='受检单位')
parser.add_argument('--type', help='型号规格')



args = parser.parse_args()
print(args.path)
path = args.path
product = args.product
samplenumber = args.samplenumber
institution = args.institution
type = args.type

##def generate_report(path, product, samplenumber, institution):
files = []
for file in os.listdir(args.path):
    if file.endswith(".txt"):
        files.append(os.path.join(args.path, file))
files.sort()


fname0 = files[0]
print(fname0)
with open(fname0, 'r') as file0:
    lines0 = file0.readlines()
    lines0=[s.strip() for s in lines0]
    time_stamp = lines0[4][6:]
    num_terminal= len(range(12, len(lines0),2))

base_data=read_data_from_file(fname0)[1]
data=np.zeros(shape=(0, 3))
delta=np.zeros(shape=(0, 3))
datetimes=[]
for fname in files:
    read_data=read_data_from_file(fname)
    new_delta=read_data[1]-base_data
    data = np.vstack((data, read_data[1]))
    delta = np.vstack((delta, new_delta))
    datetimes.append(read_data[0])


   
# plotting
length=int(len(delta)/4) 

# x是时间戳 小时
x = [(item-datetimes[0]).total_seconds()/3600.0 for item in datetimes ]

from matplotlib.font_manager import FontProperties
font = FontProperties(fname=r"simsun.ttc", size=12)


for i in range(num_terminal):
    plt.figure(1, figsize=(16, 4))
    plt.suptitle(u'端口' + str(i + 1) + '衰减变化量', fontproperties=font)
    
    ax = plt.subplot(131)
    ax.plot(x, delta[i::num_terminal, 0])
    ax.set_title('1310nm')
    ax.set_ylim(-0.5, 0.5)
    ax.set_xlabel(r'时间 (h)', fontproperties=font)
    ax.set_ylabel(r'变化量 (dB)', fontproperties=font)
    
    ax = plt.subplot(132)
    ax.plot(x, delta[i::num_terminal, 1])  
    ax.set_title('1490nm')
    ax.set_ylim(-0.5, 0.5)
    ax.set_xlabel(r'时间 (h)', fontproperties=font)
    #ax.set_ylabel(r'变化量 (dB)', fontproperties=font)
   
    ax = plt.subplot(133)
    ax.plot(x, delta[i::num_terminal, 2])
    ax.set_ylim(-0.5, 0.5)
    ax.set_title('1550nm')
    ax.set_xlabel(r'时间 (h)', fontproperties=font)
    #ax.set_ylabel(r'变化量 (dB)', fontproperties=font)
   
   
    plt.tight_layout(pad=3, w_pad=2, h_pad=4)
    plt.savefig(str(i) + '.jpg')
    
    plt.show()
    
#生成报告
start_time_string = format_datetime(datetimes[0], locale='zh_CN')
end_time_string = format_datetime(datetimes[-1], locale='zh_CN')

report = Document()


title = report.add_paragraph('试验名称：XXX试验插入损耗变化量在线监测')
title.alignment = WD_ALIGN_PARAGRAPH.CENTER


run = title.add_run()
run.font.name = "楷体_GB2312"
r = run._element
r.rPr.rFonts.set(qn('w:eastAsia'), '楷体_GB2312')




table = report.add_table(rows=4,  cols=6)
table.cell(0, 0).text = '产品名称'
table.cell(0, 1).text = 'product'
table.cell(0, 2).text = '样品编号'
table.cell(0, 3).text = 'samplenumber'
table.cell(0, 4).text = '型号规格'
table.cell(0, 5).text = 'type'

table.cell(1, 0).text = '受检单位'
table.cell(1, 1).text = 'institution'
table.cell(1, 2).text = '设备名称'
table.cell(1, 3).text = '多通道免缠绕插回损测试仪（单模）MS08B'
table.cell(1, 4).text = '出厂编号'
table.cell(1, 5).text = '1538556'
table.cell(2, 0).text = '检验时间'
table.cell(3, 0).text = '检验人员'
time_cell = table.cell(2, 1).merge(table.cell(2,2)).merge(table.cell(2,3)).merge(table.cell(2,4)).merge(table.cell(2,5))
table.cell(3, 1).merge(table.cell(3,2)).merge(table.cell(3,3)).merge(table.cell(3,4)).merge(table.cell(3,5))
time_cell.text = start_time_string + ' 至 ' + end_time_string

report.add_paragraph(' ')

for i in range(num_terminal):
    report.add_picture(str(i) + '.jpg', width=Inches(6), height=Inches(1.5))

report.save('报告.docx')
##generate_report(args.path, args.product, args.samplenumber, args.institution)