import os
import smtplib
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Using Netease 163 SMTP service
mail_host = "smtp.163.com"  # SMTP server
mail_user = "zeqili2001@163.com"  # Username
mail_pass = "SLOVWAGHQRSWOEEO"  # Authorization password, not the login password

sender = 'zeqili2001@163.com'  # Sender's email address (preferably full address, otherwise it may fail)
receivers = ['zeqili2001@163.com']  # List of recipient email addresses

title = '测试报告'  # Email subject

attachment_folder = r'C:\\Users\\ROG\\Desktop\\autorepo'  # Folder containing the attachments

def create_html_table():
    # 创建表格数据
    table_data = [
        ['姓名', '年龄', '性别'],
        ['John', '30', '男'],
        ['Alice', '25', '女'],
        ['Bob', '28', '男']
    ]

    # 设置表格样式
    table_style = """
    <style>
    table {
      border-collapse: collapse;
      width: 100%;
    }

    th, td {
      border: 1px solid black;
      padding: 8px;
      text-align: center;
    }

    th {
      background-color: #808080;
      color: white;
    }

    td {
      background-color: #E0E0E0;
    }
    </style>
    """

    # 构建HTML表格
    table_html = "<table>"
    for row in table_data:
        table_html += "<tr>"
        for cell in row:
            table_html += f"<td>{cell}</td>"
        table_html += "</tr>"
    table_html += "</table>"

    return table_style + table_html

def send_email():
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = ",".join(receivers)
    message['Subject'] = title

    # Create HTML table
    table_content = create_html_table()
    content = '来自自动脚本：\n' + table_content

    # Add email content
    message.attach(MIMEText(content, 'html'))

    # Get a list of files in the folder
    attachment_files = os.listdir(attachment_folder)

    # Add attachments
    for filename in attachment_files:
        attachment_path = os.path.join(attachment_folder, filename)
        with open(attachment_path, 'rb') as attachment_file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_file.read())

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(part)

    try:
        smtpObj = smtplib.SMTP_SSL(mail_host, 465)  # Enable SSL for secure connection, usually on port 465
        smtpObj.login(mail_user, mail_pass)  # Login authentication
        smtpObj.sendmail(sender, receivers, message.as_string())  # Send email
        print("Mail has been sent successfully.")
    except smtplib.SMTPException as e:
        print(e)

if __name__ == '__main__':
    send_email()
