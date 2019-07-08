# _*_ coding:utf-8 _*_

"""
Author:sea85
Data:2019/7/2
将srt字幕文件转换为PDF
"""

import re
from docx import Document
from docx.shared import Inches,Pt

def isdialog(str,row=0):
    res = re.search(r'Dialogue: 0.*,0,0,0,,(.*)\\N{.*}(.*)',str)
    if res:
        if not re.search(r'[\{\}]',res.group(1)):
            #print(res.groups())
            print(res.group(1))
            print(res.group(2))
            document.add_paragraph(res.group(2))
            document.add_paragraph(res.group(1))
            document.add_paragraph('\n')


path = "C:\\Users\Administrator\\Desktop\Personal\\字幕文件"
filepath = "C:\\Users\\SEAG\\Desktop\\Chernobyl.S01E01.1.23.45.720p.AMZN.WEB-DL.DDP5.1.H.264-NTb.简体&英文.ass"

document = Document()
document.add_heading('Chernobyl.S01E01', 0)

try:
    with open(filepath,'r',encoding='utf-16',errors='ignore') as f:
        for line in f:
            isdialog(line)
except UnicodeError as e:
    print(e)
    with open(filepath,'r',encoding='utf-8',errors='ignore') as f:
        for line in f:
            isdialog(line)

document.save('demo.docx')