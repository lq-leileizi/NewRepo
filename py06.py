import re
import requests
import pandas as pd
from openpyxl import Workbook

def main(num):
    global it
    url = 'http://bang.dangdang.com/books/fivestars/01.00.00.00.00.00-recent30-0-0-1-' + str(num)
    headers = {
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36"
    }
    html = requests.get(url,headers=headers)
    page =html.text

#解析数据
    pattern = re.compile(
        '<li>.*?<div class="list_num red">(?P<number>.*?).</div>',
        re.S)
#开始匹配
    result =pattern.finditer(page)
    for it in result:
        print(it.group("number"))

# 解析数据
    pattern1 = re.compile(
        '<li>.*?<div class="list_num ">(?P<number1>.*?).</div>',
        re.S)
#开始匹配
    result1 =pattern1.finditer(page)
    for it1 in result1:
        print(it1.group("number1"))


if __name__ == "__main__":
    for i in range(1,16):
        main(i)