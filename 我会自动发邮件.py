# 2023/8/1 更新 现在可以自动压缩测试文件再进行发送
# 自动发邮件脚本 V1.0.1 
# by Martin Li

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import numpy as np
import zipfile

def zip_xml_files(folder_path, zip_file_name):
    # 构建压缩文件的完整路径
    zip_file_path = os.path.join(folder_path, zip_file_name)

    # 创建一个新的zip文件
    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 遍历文件夹中的所有文件
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith('.xml'):
                    # 构建文件的完整路径
                    file_path = os.path.join(root, file)
                    # 将文件添加到zip压缩文件中
                    zipf.write(file_path, os.path.relpath(file_path, folder_path))

    print(f"XML files in '{folder_path}' have been zipped to '{zip_file_path}'.")

# Using Netease 163 SMTP service
mail_host = "smtp.163.com"  # SMTP server
mail_user = "zeqili2001@163.com"  # Username
mail_pass = "ZOIYZCITGVARKVRQ"  # Authorization password, not the login password

sender = 'zeqili2001@163.com'  # Sender's email address (preferably full address, otherwise it may fail)
receivers = ['wangmanli1@gacrnd.com']  # List of recipient email addresses

title = '版本测试报告'  # Email subject

attachment_folder = r'C:\\Users\\ROG\\Desktop\\autorepo'  # Folder containing the attachments
file_path = r'C:\\Users\\ROG\\Desktop\\autorepo\\test_results.xlsx'  # Path to the Excel file
tel = 'Tel: +86 13250231309'  # Tel number
zip_name = 'test_file.zip'  # Name of the zip file

zip_xml_files(attachment_folder, zip_name)

def create_html_table(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 获取最后六列非空列
    last_six_columns = df.iloc[:, -6:].dropna(how='all', axis=1)
    
    # 如果没有非空列，则返回空字符串
    if last_six_columns.empty:
        return ""
    
    # 将NaN值和换行符替换为空白单元格
    last_six_columns = last_six_columns.replace({np.nan: '', '\n': ''})
    
    # 生成HTML表格
    html_table = last_six_columns.to_html(index=False)
    
    print("done")
    return html_table

def send_email():
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = ",".join(receivers)
    message['Subject'] = title

    # Create HTML table
    table_content = '各位好，这个是版本测试报告，请查收 <br><br>' + '测试结果如下 <br>' + create_html_table(file_path) + '<br>' + 'Best regards, <br>' + 'Zeqi Li 黎泽麒 <br>' + '广汽研究院 智能网联技术研发中心 智驾技术部<br>' + tel + '|' + sender
    content = '来自自动脚本：<br>' + table_content

    # Add email content
    message.attach(MIMEText(content, 'html'))

    # Get a list of files in the folder
    attachment_files = os.listdir(attachment_folder)

    # Add attachments
    for filename in attachment_files:
        attachment_path = os.path.join(attachment_folder, filename)
        if not filename.endswith('.xml'):
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
for file in file_path:
    if file.endswith('.zip'):
        os.remove(file)