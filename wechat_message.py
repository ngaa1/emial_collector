import requests
from datetime import datetime
import os

def get_access_token(corpid, corpsecret):
    """
    获取企业微信 access_token
    :param corpid: 企业ID
    :param corpsecret: 应用Secret
    :return: access_token
    """
    try:
        url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={corpid}&corpsecret={corpsecret}"
        response = requests.get(url)
        result = response.json()
        if result.get('errcode') == 0:
            return result.get('access_token')
        else:
            print(f"获取 access_token 失败：{result.get('errmsg')}")
            return None
    except Exception as e:
        print(f"获取 access_token 时出错：{str(e)}")
        return None

def get_head_content():
    """
    获取head.txt文件的内容
    """
    try:
        head_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "head.txt")
        if os.path.exists(head_file):
            with open(head_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
    except Exception as e:
        print(f"读取head.txt时出错：{str(e)}")
    return ""

def send_wechat_message(corpid, corpsecret, agentid, touser, emails):
    """
    发送邮件数据到企业微信
    :param corpid: 企业ID
    :param corpsecret: 应用Secret
    :param agentid: 应用ID
    :param touser: 接收人（userid列表，用|分隔）
    :param emails: 邮件列表
    :return: 发送结果
    """
    try:
        # 获取 access_token
        access_token = get_access_token(corpid, corpsecret)
        if not access_token:
            return False
        
        # 获取head.txt的内容
        head_content = get_head_content()
        
        # 构建消息内容
        if not emails:
            content = head_content + "\n没有新邮件"
        else:
            content = head_content + f"\n共收到 {len(emails)} 封邮件\n\n"
            for i, msg in enumerate(emails, 1):
                try:
                    sender = msg.SenderName
                    subject = msg.Subject
                    received_time = msg.ReceivedTime
                    body_preview = msg.Body[:100] + "..." if len(msg.Body) > 100 else msg.Body
                    content += f"{i}. 发件人: {sender}\n"
                    content += f"   主题: {subject}\n"
                    content += f"   时间: {received_time}\n"
                    content += f"   内容: {body_preview}\n\n"
                except Exception as e:
                    print(f"处理邮件 {i} 时出错：{str(e)}")
        
        # 发送消息给企业微信app的接收消息服务器
        # 这里使用touser参数作为接收消息服务器的标识，实际上是发送给应用本身
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
        data = {
            "touser": touser,  # 这里的touser可以是任意有效的userid，或者使用@all发送给所有用户
            "msgtype": "text",
            "agentid": agentid,
            "text": {
                "content": content
            },
            "safe": 0
        }
        
        response = requests.post(url, json=data)
        result = response.json()
        if result.get('errcode') == 0:
            print("企业微信消息发送成功")
            return True
        else:
            print(f"企业微信消息发送失败：{result.get('errmsg')}")
            return False
    except Exception as e:
        print(f"发送企业微信消息时出错：{str(e)}")
        return False
