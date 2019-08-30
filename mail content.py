# _*_ coding: utf-8 _*_
import poplib
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from docx import Document
from pyquery import PyQuery as pq
import pandas as pd

def getScores(path):
    scores = {}
    document = Document(path)  # 读入文件
    tables = document.tables  # 获取文件中的表格集
    table = tables[0]  # 获取文件中的第一个表格
    for i in range(6, 19):  # 从表格第二行开始循环读取表格数据
        q = table.cell(i, 4).text
        a = table.cell(i, 5).text
        b = table.cell(i, 6).text
        c = table.cell(i, 7).text
        d = table.cell(i, 8).text
        e = table.cell(i, 9).text
        res = {'a': a, 'b': b, 'c': c, 'd': d, 'e': e}
        for k, v in res.items():
            if v:
                score = {q:k}
                scores.update(score)
    return scores

def getPath(rootdir):
    pathlist = []
    Filename  =[]
    for parent,dirnames,filenames in os.walk(rootdir):
        for filename in filenames:
            dizhi = os.path.join(parent,filename)
            pathlist.append(dizhi)
            f = filename.split(('.'))[0]
            Filename.append(f)
    return pathlist,Filename

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    return value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


def get_email_text(part):
    contentType = part.get_content_type()
    if contentType == 'text/plain' or contentType == 'text/html':
        data = part.get_payload(decode=True)
        charset = guess_charset(part)
        if charset:
            charset = charset.strip().split(';')[0]
            data = data.decode(charset)
        try:
            datas = pq(data)
            result = datas.text()
            return  result
        except:
            result = 'none'
            return result


def get_email_content(message, savepath):
    attachments = []
    for part in message.walk():
        text = get_email_text(part)
        filename = part.get_filename()
        if filename:
            filename = decode_str(filename)
            data = part.get_payload(decode=True)
            abs_filename = os.path.join(savepath, filename)
            attach = open(abs_filename, 'wb')
            attachments.append(filename)
            attach.write(data)
            attach.close()
    return attachments,text

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

if __name__ == '__main__':
    email = 'cufefinance183@163.com'
    password = 'finance183'
    pop3_server = 'pop.163.com'
    server = poplib.POP3(pop3_server)
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
        # t = int(headers['Date'].split(' ')[1])
        # if t > nowdate-2:
        attachments,text = get_email_content(msg, r'C:\Users\Administrator\Desktop\xuewei')
        with open('mailinfo.text','a+',encoding='utf-8') as f:
            content = str(text) +'subject:'+str(headers['Subject'])+'from:'+str(headers['From'])+'to:'+ str(headers['To'])+'date:'+str(headers['Date'])+'\n'
            f.write(content)
    server.quit()

