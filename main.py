import smtplib
from email.message import EmailMessage
from email.headerregistry import Address, Group
import email.policy
import mimetypes
import base64
import pandas
import time
import traceback
import openpyxl

class SendEmail(object):
    """
    python3 发送邮件类
    格式： html
    可发送多个附件
    """
    def __init__(self, smtp_server, smtp_user, smtp_passwd, sender, recipient):
        # 发送邮件服务器,常用smtp.163.com
        self.smtp_server = smtp_server
        # 发送邮件的账号
        self.smtp_user = smtp_user
        # 发送账号的客户端授权码
        self.smtp_passwd = smtp_passwd
        # 发件人
        self.sender = sender
        # 收件人
        self.recipient = recipient

        # Use utf-8 encoding for headers. SMTP servers must support the SMTPUTF8 extension
        # https://docs.python.org/3.6/library/email.policy.html
        self.msg = EmailMessage(email.policy.SMTPUTF8)

    def set_header_content(self, subject, content):
        """
        设置邮件头和内容
        :param subject: 邮件题头
        :param content: 邮件内容
        :return:
        """
        self.msg['From'] = self.sender
        self.msg['To'] = self.recipient
        self.msg['Subject'] = subject
        self.msg.set_content(content, subtype="html")

    def set_accessories(self, path_list: list):
        """
        添加附件
        :param path_list: [{"path": ""}, {"name": ""}]
        :return:
        """
        for path_dict in path_list:
            path = path_dict['path']
            name = path_dict['name']
            # print(path, name)
            ctype, encoding = mimetypes.guess_type(path)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'

            maintype, subtype = ctype.split('/', 1)
            with open(path, 'rb') as fp:
                self.msg.add_attachment(fp.read(), maintype, subtype, filename=name)
                # self.msg.add_attachment(fp.read(), maintype, subtype, filename=self.dd_b64(name))

    def send_email(self):
        """
        发送邮件
        :return:
        """
        with smtplib.SMTP_SSL(self.smtp_server, port=465) as smtp:
            # HELO向服务器标志用户身份
            smtp.ehlo_or_helo_if_needed()
            # 登录邮箱服务器
            smtp.login(self.smtp_user, self.smtp_passwd)
            print("Email:{}==>{}".format(self.sender, self.recipient))
            smtp.send_message(self.msg)
            print("成功发送!")
            smtp.quit()

    @staticmethod
    def dd_b64(param):
        """
        对邮件header及附件的文件名进行两次base64编码，防止outlook中乱码。
        email库源码中先对邮件进行一次base64解码然后组装邮件
        :param param: 需要防止乱码的参数
        :return:
        """
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        return param


if __name__ == '__main__':
    # 备用多个邮箱，抛出异常随时切换
    # 163邮箱免费版每日限流限邮件数量
    # qq邮箱时而意外断开连接，需要等1小时左右恢复
    # 谷歌邮箱或者其他邮箱应该是可以的，也有可能出现什么问题
    use_choice = 2
    if use_choice == 1:
        smtp_server = "smtp.163.com"
        smtp_user = "example1@163.com"  # 发送邮件的账号
        smtp_passwd = "POP3orSMTPcode1"  # 发送账号的客户端授权码
        sender = Address("张三", "example1", "163.com")
    elif use_choice == 2:
        smtp_server = "smtp.qq.com"
        smtp_user = "example2@qq.com"  # 发送邮件的账号
        smtp_passwd = "POP3orSMTPcode2"  # 发送账号的客户端授权码
        sender = Address("张三", "example2", "qq.com")
    elif use_choice == 3:
        smtp_server = "smtp.gmail.com"
        smtp_user = "example3@gmail.com"  # 发送邮件的账号
        smtp_passwd = "POP3orSMTPcode3"  # 发送账号的客户端授权码
        sender = Address("张三", "example3", "gmail.com")
    # 日期
    t = time.localtime()
    content_pass = f"""
        <html>
            <p>某某同学：</p>
            <p style="text-indent:2em;">  你好，请查收作业批改情况</p>

 
            <p style="text-align: right">张三</p>
            <p style="text-align: right">{t.tm_year}年{t.tm_mon}月{t.tm_mday}日</p>
        </html>
    """
    numb = '一'  # 作业批次
    subject = f'第{numb}章作业批改回复'


    df = pandas.read_excel(f'第{numb}章作业.xlsx')
    # df = pandas.read_excel(f'测试.xlsx')  # 用于测试的，可以自制表格，先发给自己的其他邮箱，检验效果进行调试
    # 获取发送次数数据 先打开我们的目标表格，再打开我们的目标表单
    wb = openpyxl.load_workbook(rf'第{numb}章作业.xlsx')
    ws = wb[f'第{numb}章作业（收集结果）']
    # 存储没发送人员名单
    nosend_list=[]
    # 统计还剩多少没发送
    g=0
    # 序号
    j=1
    # 发送次数统计
    send_num=[0]*len(df)
    for i in range(len(df)):
        j+=1
        stu_num = df.loc[i]['学号（必填）']
        stu_name = df.loc[i]['姓名（必填）']
        file_type = df.loc[i]['提交word或pdf（必填）']
        file_type =file_type[file_type.find('.'):]  # 获取文件后缀名
        stu_email_pre, stu_email_domain = df.loc[i]['邮箱（必填）'].split('@')

        recipient = Group(addresses=[Address(stu_name, stu_email_pre, stu_email_domain)])
        # 批改好的作业文件路径path以及以附件发送时的名称name
        path_list = [
            {"path": f"第{numb}章作业/{stu_num}_{stu_name}{file_type}",
            "name":  f"作业{numb}批改-{stu_name}{file_type}"}]
        # 若需要再添加别的文件，例如参考答案
        path_list2 = [
            {"path": f"第{numb}章作业参考答案.docx",
            "name":  f"第{numb}章作业参考答案.docx"}]
        file_list=[path_list,path_list2]
        # 邮件正文开头替换为当前同学的名字
        content = content_pass.replace('某某', stu_name)

        # 发送邮件
        sd = SendEmail(smtp_server, smtp_user, smtp_passwd, sender, recipient)  # 创建对象
        sd.set_header_content(subject, content)  # 设置题头和内容
        # 如果发送次数为0，执行发送
        # （这是因为第一次发送可能由于发送方或者接受方的问题，发送失败；而已经有其他发送成功的情况）
        # 可以避免重复发送或者漏发
        if j*(int(df['发送次数'].values[i]==0)):
           try:
               for fl in file_list:
                   sd.set_accessories(fl)  # 添加附件
           # 提交的word，但批改文件是pdf
           except FileNotFoundError:
               file_type = '.pdf'
               path_list = [
                   {"path": f"第{numb}章作业/{stu_num}_{stu_name}{file_type}",
                    "name": f"作业{numb}批改-{stu_name}{file_type}"}]
               try:
                   for fl in file_list:
                       sd.set_accessories(fl)  # 添加附件
               except:
                       print(f'没有{stu_name}的批改作业')
                       # 添加没有批改作业的同学名单
                       nosend_list.append(f'{j}' '-' f'{stu_name}')
                       g += 1  # 统计发送失败的数量
                       continue  # 继续下一位同学的循环，不执行后面的操作

           try:
               sd.send_email()  # 发送邮件
               # 发送成功，统计次数+1
               ws.cell(row=i + 2, column=7).value = df['发送次数'].values[i] + 1
           except:
               g += 1  # 统计发送失败的数量
               # 打印发送失败的同学学号与姓名
               print(f"{stu_num}{stu_name}Error")
               # 打印错误，不过目前作者遇到的都是“None”，可以考虑注释改行
               print(traceback.print_exc())
               # 若有需要，可以取消下行注释；但不建议，因为发送失败了会跳出循环
               # 而是建议继续下一个循环，先把可以发送的全发送了
               # exit()
    wb.save(rf'第{numb}章作业.xlsx')
    # 打印没有对应批改作业文件的同学
    if nosend_list:
        print('没有以下同学的批改作业文件')
        for ij in nosend_list:
            print(ij)
    # 打印还剩多少没发送，也就是执行一次后没能发送成功的数量
    print(f'还剩{g}没发送')


