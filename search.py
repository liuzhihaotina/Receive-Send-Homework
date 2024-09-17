# 该脚本用于查找未提交作业以及不在选课名单（补选）的学生名单
import win32com
import time
from docx2pdf import convert
import os
# # import glob
from pathlib import Path
import pythoncom
import smtplib
from email.message import EmailMessage
from email.headerregistry import Address, Group
import email.policy
import mimetypes
import base64
import pandas
import time
import traceback
import shutil

numb = '一'  # 作业批次
path = os.getcwd() + '/'
p = Path(path)  # 初始化构造Path对象
file_list = list(p.glob(f"第{numb}章作业/*"))
df = pandas.read_excel(f'课程点名册.xlsx')
df2 = pandas.read_excel(f'第{numb}章作业.xlsx')

print("以下同学未提交作业")
for i in range(len(df)):
    stu_name = df.loc[i]['姓名']
    # 如果有提交多份的，打印出来
    if str(file_list).count(stu_name)>1:
        print(f'{stu_name}提交了{str(file_list).count(stu_name)}份')
    # 没有该学生的作业文件
    if not str(file_list).count(stu_name):
        # pass
        print(stu_name)  # 输出未提交作业的学生姓名，并进行下一个循环
        continue
# 涉及补选课程的学生，没在点名册里，需要打印出来
print("以下同学未在选课名单")
flag=0
for k in range(len(df2)):
    subm_name=df2.loc[k]['姓名（必填）']
    for i in range(len(df)):
        stu_name = df.loc[i]['姓名']
        if stu_name==subm_name:
            flag=1   # 在点名册找到该学生的名字，跳出循环；进行下一个学生
            break
    if flag==0:
        print(subm_name)


