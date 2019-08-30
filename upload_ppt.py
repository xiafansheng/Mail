import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

classes = ['量化投资学','高级宏观','高级计量']

#根据课程判断及生成文件夹
def mkdir(path):
    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)
    else:
        pass

#发送课件
def sendnewppt_mail(file_path,subject='课件',body ='课件'):
    smtpserver = "smtp.163.com"
    port = 465
    sender = "xfs9619@163.com"
    psw = "xiafansheng9619"
    receiver = ["xfs9619@126.com",'xfs9619@163.com']
    msg = MIMEMultipart()
    msg["from"] = sender
    msg["to"] = ','.join(receiver)
    msg["subject"] = subject
    body = MIMEText(body, "plain", "utf-8")
    msg.attach(body)
    part = MIMEApplication(open(file_path, 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename=file_path)
    msg.attach(part)
    try:
        smtp = smtplib.SMTP()
        smtp.connect(smtpserver)
        smtp.login(sender, psw)
    except:
        smtp = smtplib.SMTP_SSL(smtpserver, port)
        smtp.login(sender, psw)
    smtp.sendmail(sender, receiver, msg.as_string())
    smtp.quit()
    return  '%s发送成功'%file_path

# 判断是否为新拷入课件、如果是发送并更新日志
def sendppt_makelog(path):
    logpath = '.\%s\log.txt'%path
    f = open(logpath)
    loghis = f.readlines()
    loghis = [f.strip('\n') for f in loghis]
    for parent,dirnames,filenames in os.walk(path):
        for filename in filenames:
            if filename not in loghis:
                dizhi = os.path.join(parent, filename)
                sendnewppt_mail(dizhi)
                with open(logpath,'a+',encoding='utf-8') as f:
                    content = filename + '\n'
                    f.write(content)

for c in classes:
    sendppt_makelog(c)
