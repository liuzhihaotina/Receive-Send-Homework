# 该脚本实现对收集的作业文件进行分配
# 一门课程可能有多个助教，需要分配作业批改任务，将所有的文件按序号划分到几个单独的文件夹
import win32com
import time
from docx2pdf import convert
import os
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

numb = '二'  # 作业批次
path = os.getcwd() + '/'
p = Path(path)  # 初始化构造Path对象
file_list = list(p.glob(f"优化算法作业{numb}/*"))
df = pandas.read_excel(f'作业分配名单.xlsx')
df2 = pandas.read_excel(f'优化算法作业{numb}.xlsx')
j=1
for i in range(len(df)):
    j+=1
    stu_name = df.loc[i]['姓名']
    if not str(file_list).count(stu_name):
        # pass
        print(stu_name)  # 输出未提交作业的学生姓名，并进行下一个循环
        continue
    file_c = 0
    for k in range(len(df2)):
        if stu_name == df2.loc[k]['姓名（必填）']:
            stu_num = df2.loc[k]['学号（必填）']
            file_type = df2.loc[k]['提交word或pdf文件（必填）']
            file_type = file_type[file_type.find('.'):]  # 获取文件后缀名
            file_copy = file_type # 拷贝文件后缀名
            file_c+=1

            fenpei = 1
            if fenpei:
                # 如果有同学提交了多份作业，只拷贝最新的文件
                if file_c>1:
                    file_type=f'（{file_c-1}）'+file_type
                if j<=55:
                    # shutil.copy("C:/a/2.txt", "C:/b/121.txt")  # 复制并重命名新文件
                    # print(stu_name)
                    shutil.copy(f'优化算法作业{numb}/{stu_num}_{stu_name}{file_type}', \
                                f'作业{numb}分配1/上午1-54王/{stu_num}_{stu_name}{file_copy}')
                elif j<=108:
                    shutil.copy(f'优化算法作业{numb}/{stu_num}_{stu_name}{file_type}', \
                                f'作业{numb}分配1/上午55-107薛/{stu_num}_{stu_name}{file_copy}')
                elif j<=157:
                    shutil.copy(f'优化算法作业{numb}/{stu_num}_{stu_name}{file_type}', \
                                f'作业{numb}分配1/下午1-49刘/{stu_num}_{stu_name}{file_copy}')
                elif j<=205:
                    shutil.copy(f'优化算法作业{numb}/{stu_num}_{stu_name}{file_type}', \
                                f'作业{numb}分配1/下午50-97方/{stu_num}_{stu_name}{file_copy}')
    # 如果有同学提交了多份作业，打印输出
    if file_c>1:
        print(f'{stu_name}提交了{file_c}份作业')

