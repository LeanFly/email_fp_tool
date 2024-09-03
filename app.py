# -*- encoding: utf-8 -*-
'''
@File    :   mail_down.py
@Time    :   2024/09/02 14:09:11
@Author  :   leanfly
@Version :   1.0
@Contact :   ningmuyu@live.com
@Desc    :   当前文件作用
'''

import imaplib
import email
import os
from pathlib import Path
import time

from email.utils import parsedate_to_datetime
from email.header import make_header, decode_header

from bs4 import BeautifulSoup

from icecream import ic

from urllib.request import urlretrieve
import html

from datetime import datetime



# QQ 邮箱的 IMAP 服务器地址和端口号
imap_server = 'imap.qq.com'
port = 993

# 你的 QQ 邮箱账号和 POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 开始后的密钥
username = ""
password = ""

# 发票下载目录
fp_save_dir = ""

# 发票信息监听的关键词列表
fp_keyword = ["发票PDF文件下载"]


# 连接到 IMAP 服务器
mail = imaplib.IMAP4_SSL(imap_server, port)

# 登录邮箱
mail.login(username, password)

# 选择收件箱
mail.select('INBOX')

# 搜索包含发票的邮件
typ, data = mail.search(None, 'ALL')


def demo2():
    # 搜索邮件
    typ, data = mail.search(None, 'ALL')

    # 遍历邮件
    for num in data[0].split():
        typ, data = mail.fetch(num, '(RFC822)')
        
        if not data:
            continue
        if not data[0]:
            continue
        if not data[0][1]:
            continue
        
        raw_email = data[0][1]
        
        # 解析邮件
        msg = email.message_from_bytes(raw_email)
        
        # 获取邮件主题
        # subject = decode_header(msg['Subject'])[0][0]
        # if isinstance(subject, bytes):
        #     subject = subject.decode('utf-8', errors='replace')
        
        subject = make_header(decode_header(msg['Subject']))
        print("邮件主题:", (subject))
        if "发票" not in str(subject):
            # 邮件主题不包含发票，跳过
            continue
        
        # 获取发件人
        # from_ = decode_header(msg['From'])[0][0]
        # if isinstance(from_, bytes):
        #     from_ = from_.decode()

        from_ = make_header(decode_header(msg["From"]))
        print("发件人:", from_)
        # 京东JD.com <customer_service@jd.com>
        if "京东JD.com <customer_service@jd.com>" not in str(from_):
            # 邮件发送人不是京东客服，跳过
            continue
        
        # date_ = parsedate_to_datetime(msg['Date']).strftime("%Y-%m-%d %H:%M:%S") if msg['Date'] else ""  # 收件时间 
        date_str = ""
        date_timestamp = 0
        if msg["Date"]:
            # 这里是预留邮件时间判断逻辑，比如只处理某个时间节点之后的邮件
            date_obj = parsedate_to_datetime(msg['Date'])
            date_str = date_obj.strftime("%Y-%m-%d %H:%M:%S")
            date_timestamp = date_obj.timestamp()
        
        
        # 获取正文
        if msg.is_multipart():
            for part in msg.walk():
                # ic(dir(part))
                # part.set_charset("utf-8")
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition:
                    # ic(part.get_payload(decode=True))
                    body = part.get_payload(decode=True)
                    if not body:
                        continue
                    # body = part.get_payload(decode=True).decode()
                    body = part.get_payload(decode=True).decode("ISO-8859-1").encode("utf-8")   # 增加了解码和再编码，避免出现打印的内容是乱码

                    ic(f"Subject: {subject}, From: {from_}, Date: {date_str}")
                    # print('Body:', BeautifulSoup(body, "html.parser").find_all("a"))
                    body_links = BeautifulSoup(body, "html.parser").find_all("a")
                    for a in body_links:
                        # title = a.text
                        if a.text in fp_keyword:
                            url = a.get("href")
                            url = html.unescape(url)
                            pdf_name = Path(url).stem + "_" + datetime.now().strftime("%Y-%m-%d_%H:%M:%S") + ".pdf"
                            try:
                                urlretrieve(url, Path(fp_save_dir) / pdf_name)
                            except Exception as e:
                                ic(f"下载发票异常：{str(e)}")
                    time.sleep(1)
        else:
            # body = msg.get_payload(decode=True).decode()
            body = msg.get_payload(decode=True).decode()   # 增加了解码和再编码，避免出现打印的内容是乱码

            ic(f"Subject: {subject}, From: {from_}, Date: {date_str}")
            
            body_links = BeautifulSoup(body, "html.parser").find_all("a")
            for a in body_links:
                # title = a.text
                if a.text in fp_keyword:
                    url = a.get("href")
                    url = html.unescape(url)
                    pdf_name = Path(url).stem + "_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".pdf"
                    try:
                        urlretrieve(url, Path(fp_save_dir) / pdf_name)
                    except Exception as e:
                        ic(f"下载发票异常：{str(e)}")
            time.sleep(1)


    # 关闭连接
    mail.close()
    mail.logout()   


if __name__ == "__main__":
    
    demo2()
