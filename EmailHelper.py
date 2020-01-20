#!/usr/bin/env python3
# -*-coding:utf-8-*-

from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import parseaddr, formataddr
import smtplib
import datetime

class Email(object):

    def __init__(self):
        '''
        初始化
        '''
        # 配置发件人和发件邮箱
        self.from_addr = {'name':'发件人', 'email':'邮箱'}
        # 邮箱授权码,注意这里不是邮箱密码
        self.password = '***'
        # SMTP服务器，这里默认为 163 的
        self.smtp_server = 'smtp.163.com'


        # 构建MIMEMultipart对象代表邮件本身，可以往里面添加文本、图片、附件等
        self.mm = MIMEMultipart('related')
        # 设置发送人
        self.mm['From'] = self._format_addr('{name} <{email}>'.format(name = self.from_addr.get('name'), email = self.from_addr.get('email')))
        # 接收人
        self.receivers = []
        self.content = '<html>{header}<body>'
        self.imageid = 1

    def _format_addr(self, s):
        '''
        格式化发件人格式
        :return:
        '''
        name, addr = parseaddr(s)
        return formataddr((Header(name, 'utf-8').encode(), addr))

    def add_header(self, header=''):
        '''
        添加 header 内容，可以用来设置 html 的样式
        '''
        self.content = self.content.format(header=header)

    def add_receivers(self, receivers):
        '''
        添加收件人 显示
        :param receivers: 收件人字典 {'name1':'1213@qq.com', 'name2':'1314@qq.com'}
        :return:
        '''
        to_addr = []
        for name, email in receivers.items():
            to_addr.append(self._format_addr('{name} <{email}>'.format(name=name, email=email)))
            self.receivers.append(email)

        to_addr_str = ",".join(to_addr)
        self.mm['To'] = to_addr_str

    def add_copy(self, copy={}):
        '''
        添加抄送人 显示
        :param copy: 抄送人字典 {'ding':'1213@qq.com'}
        :return:
        '''
        copy_addr = []
        for name, email in copy.items():
            copy_addr.append(self._format_addr('{name} <{email}>'.format(name=name, email=email)))
            self.receivers.append(email)

        copy_addr_str = ",".join(copy_addr)
        self.mm['Cc'] = copy_addr_str

    def add_subject(self, subject=''):
        '''
        添加主题
        :param subject: 主题文本
        :return:
        '''
        # 设置邮件主题
        self.mm["Subject"] = Header(subject, 'utf-8')

    def add_text(self, text=''):
        '''
        添加正文
        :param body: 正文内容
        :return:
        '''
        self.content += text

    def add_img(self, path):
        '''
        添加图片
        :param path: 图片地址
        :return:
        '''
        imageid = self.imageid
        # 二进制读取图片
        image_data = open(path, 'rb')
        # 设置读取获取的二进制数据
        message_image = MIMEImage(image_data.read())
        # 关闭刚才打开的文件
        image_data.close()
        message_image.add_header('Content-ID', str(imageid))
        # 添加图片文件到邮件信息当中去
        self.mm.attach(message_image)
        self.content += '<br/><img src="cid:{imageid}" alt="{imageid}">'.format(imageid=str(imageid))
        self.imageid +=1

    def add_attachment(self, path, filename):
        '''
        添加附件
        :param path: 附件的路径
        :param filename: 添加附件的名称
        :return:
        '''
        # 构造附件
        atta = MIMEText(open(path, 'rb').read(), 'base64', 'utf-8')
        # 设置附件信息
        atta["Content-Disposition"] = 'attachment; filename="{filename}"'.format(filename=filename)
        # 添加附件到邮件信息当中去
        self.mm.attach(atta)

    def send_email(self):
        '''
        发送
        :return:
        '''
        # 内容结尾
        self.content += '</body></html>'
        # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
        content = MIMEText(self.content, 'html', 'utf-8')
        self.mm.attach(content)

        # 创建SMTP对象
        stp = smtplib.SMTP()
        # 设置发件人邮箱的域名和端口，端口地址为25
        stp.connect(self.smtp_server, 25)
        # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
        stp.set_debuglevel(1)
        # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
        stp.login(self.from_addr.get('email'), self.password)
        # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
        print(self.receivers)
        stp.sendmail(self.from_addr.get('email'), self.receivers, self.mm.as_string())
        print("邮件发送成功")
        # 关闭SMTP对象
        stp.quit()



if __name__ == '__main__':
    email = Email()
    # 不需要的时候填写为
    header = ''

    email.add_header(header)
    # 添加收件人, 添加上发送人，防止被过滤为垃圾邮件
    receivers = {'name1':'******@qq.com', 'name2':'******@163.com', 'name3': '*****@163.com'}
    email.add_receivers(receivers)
    # 添加抄送人
    copy = {'name1':'*******@qq.com', 'name2':'******@163.com'}
    email.add_copy(copy)
    # 添加主题
    subject = '日报邮件'
    email.add_subject(subject)
    # 添加正文内容
    text = '明天来加班'
    email.add_text(text)
    # 添加图片
    email.add_img('test.jpg')
    # 添加 html 表格
    table = '''
    <body>
	<table border="1" cellpadding="0" cellspacing="0" width="400px" style="border-collapse: collapse;">
		<tbody>
			<tr style="color:white;font-weight:700;font-family:微软雅黑;background:black;">
				<th>主站销售额</th>
				<th></th>
			</tr>
			<tr style="background:#BFBFBF;">
				<td style="text-align:center">2019年11月</td>
				<td style="text-align:center">更新日期：11月16日</td>
			</tr>
		</tbody>
    </table> 
    '''
    email.add_text(table)

    file_time = datetime.datetime.strptime('2019-2-16 00:00:00', '%Y-%m-%d %H:%M:%S')
    filename = file_time.strftime("%Y{y}%m{m}%d{d}").format(y='年', m='月', d='日')
    # 上传附件名称，不要有中文
    en_filename = file_time.strftime("%Y-%m-%d")
    print('./{filename}.xlsx'.format(filename=filename), filename + '.xlsx')
    # 添加附件的路径, 邮件中附件显示的名称(邮件中显示的名称不要用中文)
    email.add_attachment('./{filename}.xlsx'.format(filename=filename), en_filename+'.xlsx')

    email.send_email()