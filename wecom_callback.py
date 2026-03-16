import requests
from datetime import datetime
import os


def send_to_wecom_callback(callback_url, token, encoding_aes_key, emails):
    """
    发送邮件数据到企业微信接收消息服务器
    :param callback_url: 接收消息服务器URL
    :param token: Token
    :param encoding_aes_key: EncodingAESKey
    :param emails: 邮件列表
    :return: 发送结果
    """
    try:
        # 构建邮件数据
        email_data = []
        for msg in emails:
            try:
                email_info = {
                    "sender": msg.SenderName,
                    "subject": msg.Subject,
                    "received_time": str(msg.ReceivedTime),
                    "body": msg.Body
                }
                email_data.append(email_info)
            except Exception as e:
                print(f"处理邮件时出错：{str(e)}")
        
        # 构建请求数据
        data = {
            "email_count": len(emails),
            "emails": email_data,
            "timestamp": datetime.now().isoformat(),
            "token": token,
            "encoding_aes_key": encoding_aes_key
        }
        
        # 发送POST请求
        response = requests.post(callback_url, json=data, headers={"Content-Type": "application/json"})
        
        if response.status_code == 200:
            print("邮件数据发送到企业微信接收消息服务器成功")
            return True
        else:
            print(f"邮件数据发送到企业微信接收消息服务器失败：{response.status_code} - {response.text}")
            return False
    except Exception as e:
        print(f"发送到企业微信接收消息服务器时出错：{str(e)}")
        return False
