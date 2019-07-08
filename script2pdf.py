# _*_ coding:utf-8 _*_

"""
Author:sea85
Data:2019/7/2
将srt字幕文件转换为PDF
"""

import re

def isdialog(str,row=0):
    res = re.search(r'Dialogue: 0.*NTP,0,0,0,,(.*)\\N{.*}(.*)',str)
    if res:
        print(res.group(1))
        print(res.group(2))

path = "C:\\Users\Administrator\\Desktop\Personal\\字幕文件"
filepath = "C:\\Users\Administrator\\Desktop\Personal\\字幕文件\\Game.of.Thrones.S08E01.Kings.Landing.720p.AMZN.WEB-DL.DDP5.1.H.264-GoT.简体&英文.ass"

try:
    with open(filepath,'r',encoding='utf-16',errors='ignore') as f:
        for line in f:
            isdialog(line)
except UnicodeError as e:
    print(e)
    with open(filepath,'r',encoding='utf-8',errors='ignore') as f:
        for line in f:
            isdialog(line)