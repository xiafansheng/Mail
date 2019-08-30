# _*_ coding: utf-8 _*_
import poplib
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from docx import Document
import pandas as pd


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
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
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
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
            filename = decode_str(filename)
            data = part.get_payload(decode=True)
            abs_filename = os.path.join(savepath, filename)
            attach = open(abs_filename, 'wb')
            attachments.append(filename)
            attach.write(data)
            attach.close()
    return attachments


if __name__ == '__main__':
    nowdate = int(input('请输入今天是几号'))
    email ='2467028805@qq.com'
    password ='dpouackbvbpsebfc'
    pop3_server ='pop.qq.com'
    server = poplib.POP3_SSL(pop3_server,995)
    server.set_debuglevel(0)
    server.user(email)
    server.pass_(password)
    msg_count, msg_size = server.stat()
    resp, mails, octets = server.list()
    for i in range(msg_count,1,-1):
        resp, byte_lines, octets = server.retr(i)
        str_lines = []
        for x in byte_lines:
            str_lines.append(x.decode())
        msg_content = '\n'.join(str_lines)
        msg = Parser().parsestr(msg_content)
        headers = get_email_headers(msg)
        t = int(headers['Date'].split(' ')[1])
        if t > nowdate-3:
            attachments = get_email_content(msg, r'C:\Users\Administrator\Desktop\attachment')
            print('subject:', headers['Subject'])
            print('from:', headers['From'])
            print('to:', headers['To'])
            if 'cc' in headers:
                print('cc:', headers['Cc'])
            print('date:', headers['Date'])
            print('attachments: ', attachments)
            print('-----------------------------')
        else:
            break


