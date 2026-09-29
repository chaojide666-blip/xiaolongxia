# -*- coding: utf-8 -*-
"""
饭加鱼：邮件技能 v1
支持发送纯文本邮件
"""

import smtplib
from email.mime.text import MIMEText
from email.header import Header


def send_email(sender_email, auth_code, receiver_email, subject, body, smtp_server, smtp_port=465):
    """
    发送邮件。
    
    参数：
    - sender_email: 发件人邮箱
    - auth_code: 邮箱授权码（不是登录密码）
    - receiver_email: 收件人邮箱
    - subject: 邮件主题
    - body: 邮件正文
    - smtp_server: SMTP服务器地址
    - smtp_port: SMTP端口，默认465（SSL）
    """
    if not sender_email or not auth_code or not receiver_email:
        return {"success": False, "error": "发件人、授权码或收件人不能为空。"}

    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"] = sender_email
        msg["To"] = receiver_email

        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()

        server.login(sender_email, auth_code)
        server.sendmail(sender_email, [receiver_email], msg.as_string())
        server.quit()

        return {
            "success": True,
            "message": f"邮件已发送给 {receiver_email}",
        }

    except Exception as e:
        return {
            "success": False,
            "error": f"邮件发送失败：{e}",
        }