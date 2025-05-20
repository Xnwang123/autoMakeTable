import smtplib
from email.mime.text import MIMEText

server = smtplib.SMTP_SSL('smtp.qq.com',465)
server.login("2696572657@qq.com","ggstbxmdxiuadfca")
msg = "中国"
message = MIMEText(msg,'plain','utf8')
message["From"] = "{}".format("2696572657@qq.com")
message["to"] = ",".join(['703277461@qq.com'])
server.sendmail("2696572657@qq.com","703277461@qq.com",message.as_string())

server.quit()