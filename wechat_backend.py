import requests
from datetime import datetime

def send_to_wechat_backend(backend_url, emails):
    """
    发送邮件数据到企业微信应用后台
    :param backend_url: 应用后台URL
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
            "timestamp": datetime.now().isoformat()
        }
        
        # 发送POST请求
        response = requests.post(backend_url, json=data, headers={"Content-Type": "application/json"})
        
        if response.status_code == 200:
            print("邮件数据发送到企业微信应用后台成功")
            return True
        else:
            print(f"邮件数据发送到企业微信应用后台失败：{response.status_code} - {response.text}")
            return False
    except Exception as e:
        print(f"发送到企业微信应用后台时出错：{str(e)}")
        return False
