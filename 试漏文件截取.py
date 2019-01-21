import re
import xlwt
import os
import time

# 创建正则表达式，拆分原始文件的标签和内容
LABEL_PATTERN = re.compile("<(?P<label>\S+)>.+?</(?P=label)>")
LABEL_CONTENT_PATTERN = re.compile("<(?P<label>\S+)>(.*?)</(?P=label)>")

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
PATH 是原始XML文件的存放地址，如果你把原始文件放在了别的地址，这个地方需要改！
注意！！！
“\”需要写两次！！！
比如，你把文件都放到了 “E:\Abc\123” 这个文件夹下面，那么这里你就要写：
"E:\\Abc\\123"
看到没，每个"\"都要写两次！

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
PATH = "D:\\LT_Results"

# 这里是设定表头

list_key = ["PARTNO", "CHANNEL_NO", "EVAL", "VALUE"]
title = ["发动机号", "通道名称1", "结果", "值1", "值2", "通道名称2", "结果", "值1", "值2", "通道名称3", "结果", "值1", "值2", "通道名称4", "结果", "值1",
         "值2", ]

wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("试漏结果")
col_title = 0
for i in title:
    ws.write(0, col_title, i)
    col_title += 1

# 这里是填写数据

line_no = 1
files = os.listdir(PATH)
for file in files:
    col_no = 0

    f = open(PATH + "\\" + file)
    content = f.read()
    f.close()
    values = LABEL_CONTENT_PATTERN.findall(content)

    for i in values:
        x, y = i
        if x in list_key:
            ws.write(line_no, col_no, y)
            col_no += 1
    line_no += 1

# 这里是操作写 excel 文件，并且修改文件名
# 文件名的格式是：试漏结果 + 当前日期 + 当前时间

current_time = time.strftime("%Y_%m_%d_%X", time.localtime())
final_time = current_time.replace(":", "_")
print(final_time)
final_str = "D:\\试漏结果" + final_time + ".xls"

print(final_str)
wb.save(final_str)
