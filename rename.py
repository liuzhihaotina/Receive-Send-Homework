# 腾讯文档VIP不仅拥有100G空间，还能自定义收集作业的文件名
# 1. 但不用VIP也是可以的，这个文件一样可以按收集表的 学号_姓名 重命名文件
# 2. 此外，免费用户只有1G空间，需要及时清理之前收集的文件
import os
import pandas

dir_path = os.path.dirname(os.path.abspath(__file__))   #获取 绝对路径主目录
xlsbpath=rf"{dir_path}/作业文件夹"
os.chdir(xlsbpath) #更改当前路径
filelist = os.listdir(xlsbpath)  # 该文件夹下所有的文件（包括文件夹）
# print(filelist) #文件夹中所有文件名

e='测试.xlsx'  # 收集表导出的表格文件
df = pandas.read_excel(f'{dir_path}/{e}')
for i in range(len(df)):
    stu_num = df.loc[i]['学号（必填）']
    stu_name = df.loc[i]['姓名（必填）']
    file_name = df.loc[i]['提交word或pdf（必填）']
    file_type = file_name[file_name.find('.'):]  # 获取文件后缀名

    os.rename(file_name, f'{stu_num}_{stu_name}{file_type}')  # 重命名