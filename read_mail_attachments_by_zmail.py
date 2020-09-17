#!/user/bin/python
# _*_ coding:utf-8 _*_
import xlrd
import zmail

__author__ = "super.gyk"


class ControlMail(object):

    def __init__(self, name, pas, sender, sub, content):
        self.user_name = name
        self.user_pass = pas
        self.pop_host = 'pop.exmail.qq.com'
        self.smtp_host = 'smtp.exmail.qq.com'
        self.sender = sender
        self.sub = sub
        self.content = content
        # 如果是腾讯企业邮箱需添加pop_host和smtp_post,不填默认为普通QQ邮箱
        self.server = zmail.server(self.user_name, self.user_pass)

    # 读取邮箱附件Excel
    def read_excel(self, file_contents):
        data_list = []
        data_list1 = []
        # 此处传输接收到的是文件内容而不是文件对象，故用替换file_name
        data = xlrd.open_workbook(file_contents=file_contents)
        t = data.sheets()[0]
        n_rows = t.nrows
        n_cols = t.ncols
        for i in range(1, n_rows):
            for j in range(n_cols):
                vl = t.cell(i, j).value
                data_list1.append(vl)
            data_list.append(data_list1)
        print(data_list)
        return data_list

    # 发送邮件
    def send_mails(self):
        mail_content = {
            'subject': self.sub,
            'content_text': self.content
        }
        self.server.send_mail(self.sender, mail_content)

    # 读邮件
    def read_mail(self):
        mails = self.server.get_mails()
        for item in mails:
            print(zmail.show(item))

    # 读邮件附件
    def read_mail_attachment(self):
        mails = self.server.get_mails()
        for item in mails:
            # 判读是否有附件
            if item['Attachments']:
                print(item['Id'], item['Subject'])
                for name, raw in item['attachments']:
                    if name.split('.')[1] in ['xls', 'xlsx']:
                        self.read_excel(raw)


if __name__ == '__main__':
    user_name = '1xxx@qq.com'
    user_pass = 'riayxxxhomqiifrbabh'
    sender = '1xxxx@qq.com'
    subject = '测试'
    content = '测试内容'
    mail = ControlMail(user_name, user_pass, sender, subject, content)

    # mail.send_mails()
    # mail.read_mail()
    # mail.read_mail_attachment()
