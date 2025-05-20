import smtplib
from email.mime.text import MIMEText

isSucess = False
def feedBack(info_name,text):
    # 配置参数
    sender_email = "2696572657@qq.com"
    password = "ggstbxmdxiuadfca"     #SSL授权码
    receiver_email = "703277461@qq.com"

    # 创建邮件内容
    subject = f"来自{info_name}反馈"
    content = text
    message = MIMEText(content, 'plain', 'utf-8')
    message['From'] = "{}".format("2696572657@qq.com")
    message['To'] = ",".join(['703277461@qq.com'])
    message['Subject'] = subject

    try:
        # 建立SSL加密连接
        server = smtplib.SMTP_SSL('smtp.qq.com',465)  #设置SMTP服务器以及端口
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        isSucess = True
        print("邮件发送成功！")
    except Exception as e:
        print(f"发送失败，错误信息：{str(e)}")
        isSucess = False
    return isSucess

"""
不同邮箱服务商配置参考
邮箱类型	SMTP服务器	端口	加密方式	密码类型
Gmail	smtp.gmail.com	465	SSL	应用专用密码
QQ邮箱	smtp.qq.com	465	SSL	邮箱授权码
163邮箱	smtp.163.com	465	SSL	客户端授权密码
Outlook	smtp.office365.com	587	STARTTLS	账户密码
"""