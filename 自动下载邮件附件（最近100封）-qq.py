# -*- coding: utf-8 -*-
# coding: utf-8
# 引入模块及IMAPClient类
import email
import poplib
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import os
from datetime import datetime,date,timedelta
d=os.getcwd()

#初始化邮箱信息
hostname='smtp.qq.com'  #smtp服务器网址（注意这里面的hostname问题）
username='XXXX'  #邮件用户名
passwd='XXXX'  #邮箱密码，如果是QQ邮箱，则是授权码,使用的时候，其他访问要关闭，比如邮箱的网页端

#连接到POP3服务器，带SSL的：
server=poplib.POP3_SSL(hostname)
server.set_debuglevel(0)
print(server.getwelcome())  #POP3服务器的欢迎文字

#身份认证
server.user(username)
server.pass_(passwd)
#stat()返回邮件数量和占用空间：
msg_count,msg_size=server.stat()
print('message count:',msg_count)
print('message size',msg_size,'bytes')

def decode_str(s):
    value, charset = decode_header(s)[0]#数据,数据编码方式，from email.header import decode_header
    if charset:
        value = value.decode(charset)
    return value

def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            if header == 'Subject':
                subject = decode_str(value)
                headers['Subject'] = subject
            if header == 'From':
                hdr,addr = parseaddr(value)
                name=decode_str(hdr)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            if header == 'To':
                all_cc = value.split(',')
                to = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            if header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    hdr, addr = parseaddr(x)
                    name = decode_str(hdr)
                    cc_addr = u'%s <%s>' % (name, addr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc)
    return headers

def get_email_content(message, savepath):
    attachments = []
    for part in message.walk():
        filename = part.get_filename()
        if filename:
            filename=decode_str(filename)
            data = part.get_payload(decode=True)
            abs_filename = os.path.join(savepath, filename)
            attach = open(abs_filename, 'wb')
            attachments.append(filename)
            attach.write(data)
            attach.close()
    return attachments
today = datetime.today()  #今天日期
yesterday=today+timedelta(days=-1)  #昨天日期
data_path=d + '\\attach_qq\\'+str(today.date())+'\\' #目录的符号别忘了
if not os.path.exists(data_path):
    os.makedirs(data_path)  #创建文件夹，以日期命名文件夹
sender=[]
for i in range(msg_count-100,msg_count):
    resp,byte_lines,cotets=server.retr(i)
    #转码
    str_lines=[]
    for x in byte_lines:
        str_lines.append(x.decode('gbk'))  #这个地方很关键，必须修改为'gbk'解码，否则会报错！
    msg_content='\n'.join(str_lines)
    #把邮件内容解析为message对象
    msg=Parser().parsestr(msg_content)
    headers=get_email_headers(msg)
    attachments=get_email_content(msg,d+'\\attach_qq\\')
    #输出
    sender.append(headers['From'].split('<')[1][:-1])  #将发件人的邮箱存起来，之后可以批量回复
    # print('subject:',headers['Subject'])
    # print('from:',headers['From'])
    # print('to:',headers['To'])
    # if 'cc'in headers:
    #     print('cc',headers['Cc'])
    # print('date:',headers['Date'])
    # print('attachments:',attachments)
    # print('----------------------------------------------------------------------')
print('最近100封邮件及附件接收完毕')