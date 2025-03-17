import requests
import re
import sys
from io import StringIO
import openpyxl

def main(num):
    global it
    url = 'http://bang.dangdang.com/books/fivestars/01.00.00.00.00.00-recent30-0-0-1-' + str(num)
    headers = {
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36"
    }
    html = requests.get(url, headers=headers)
    page = html.text

    # 解析数据
    pattern = re.compile(
        '<li>.*?target="_blank" title="(?P<name>.*?)">',
        re.S)
    # 开始匹配
    result = pattern.finditer(page)
    for it in result:
        print(it.group("name"))

if __name__ == "__main__":
    # 重定向标准输出以捕获书名
    original_stdout = sys.stdout
    sys.stdout = StringIO()

    try:
        for i in range(1,16):
            main(i)
    finally:
        # 恢复标准输出
        captured_output = sys.stdout.getvalue()
        sys.stdout = original_stdout

    # 将捕获的输出分割成书名列表
    book_titles = captured_output.splitlines()

    # 创建Excel文件并保存数据
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Book Titles"
    ws.append(["书名"])  # 添加表头

    for title in book_titles:
        ws.append([title])

    wb.save("当当图书榜单.xlsx")
    print(f"数据已保存到 当当图书榜单.xlsx，共{len(book_titles)}条记录。")