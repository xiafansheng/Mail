import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pandas as pd

df = pd.read_excel(r'C:\Users\Administrator\Desktop\《零售银行》报名2.xlsx').ix[:,:5]
df.columns =  ['姓名', '手机', '公司名称', '职位', '邮箱',]
mail,name = df['邮箱'].values,df['姓名'].values
m_n = {i[0]:i[1] for i in zip(mail,name) if i[0] !='nan'}


#m_n = {'2467028805@qq.com':'夏凡盛'}

def send_inform(file_path,subject,body,filename,r):
    smtpserver = "smtp.163.com"
    port = 465
    sender = "xfs9619@163.com"
    psw = "xfs9619"
    receiver = [r]  #"cufe2018jrxs@163.com"  'xfs9619@163.com'
    msg = MIMEMultipart()
    msg["from"] = sender
    msg["to"] = ','.join(receiver)
    msg["subject"] = subject
    body = MIMEText(body, "plain", "utf-8")
    msg.attach(body)
    part = MIMEApplication(open(file_path, 'rb').read())
    part.add_header('Content-Disposition', 'attachment', filename=('gbk', '', filename) )
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
count =0

for k,v in m_n.items():
    try:
        body = '尊敬的%s 先生/女士：\n\
                您好，我是第二届中国金融科技前沿论坛的工作人员,非常感谢您报名参加此次大会，若您确定于4月27日在中央财经大学参加本次会议，请按"姓名，确定参会"回复此邮件!\n\
                说明：本次论坛不收会议费，交通食宿自理，论坛提供27日中午自助餐，需要正式邀请函的参会代表请在会议报道处领取盖有公章的邀请函。本次大会会议议程添加在附件中，望查收，期待您的参与！\n\
                第二届金融科技前沿论坛\n\
                2019年4月19日'%v
        file = r'C:\Users\Administrator\Desktop\第二届中国金融科技前沿论坛议程.pdf'
        s = '第二届中国金融科技前沿论坛报名确认'
        send_inform(file,s,body,'2019中国金融科技前沿论坛--会议议程.pdf',k)
        count+=1
        print(count,'%s发送成功'%v)
    except:
        print(v,'error')