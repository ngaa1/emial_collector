import requests
from datetime import datetime
import os
import json
from wecom_crypto import WeComCrypto


def save_request_log(log_data):
    """
    保存请求日志到文件
    :param log_data: 日志数据字典
    """
    try:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_file = os.path.join(log_dir, f"wecom_callback_{datetime.now().strftime('%Y%m%d')}.log")
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"\n{'='*80}\n")
            f.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*80}\n")
            
            for key, value in log_data.items():
                f.write(f"\n{key}:\n")
                if isinstance(value, dict):
                    for k, v in value.items():
                        f.write(f"  {k}: {v}\n")
                elif isinstance(value, list):
                    for item in value:
                        f.write(f"  - {item}\n")
                else:
                    f.write(f"  {value}\n")
        
        print(f"详细日志已保存到: {log_file}")
    except Exception as e:
        print(f"保存日志文件时出错: {str(e)}")


def send_to_wecom_callback(callback_url, token, encoding_aes_key, emails, corpid=None):
    """
    发送邮件数据到企业微信接收消息服务器
    :param callback_url: 接收消息服务器URL
    :param token: Token
    :param encoding_aes_key: EncodingAESKey
    :param emails: 邮件列表
    :param corpid: 企业ID（用于加密）
    :return: 发送结果
    """
    if corpid is None:
        corpid = ""
    
    log_data = {
        "请求开始时间": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "URL": callback_url,
        "Token": token,
        "EncodingAESKey": encoding_aes_key,
        "CorpID": corpid,
        "邮件数量": len(emails)
    }
    
    try:
        # 构建邮件数据
        email_data = []
        for i, msg in enumerate(emails, 1):
            try:
                email_info = {
                    "sender": msg.SenderName,
                    "subject": msg.Subject,
                    "received_time": str(msg.ReceivedTime),
                    "body": msg.Body[:200] + "..." if len(msg.Body) > 200 else msg.Body
                }
                email_data.append(email_info)
            except Exception as e:
                print(f"处理邮件 {i} 时出错：{str(e)}")
                log_data[f"邮件{i}错误"] = str(e)
        
        log_data["邮件数据预览"] = email_data[:2] if len(email_data) > 2 else email_data
        
        # 构建完整数据
        message_data = {
            "email_count": len(emails),
            "emails": email_data,
            "timestamp": datetime.now().isoformat()
        }
        
        print(f"\n{'='*80}")
        print(f"正在发送请求到: {callback_url}")
        print(f"原始消息数据: {message_data}")
        print(f"{'='*80}\n")
        
        log_data["原始消息数据"] = message_data
        
        # 尝试方式1：加密发送（企业微信标准格式）
        try:
            crypto = WeComCrypto(token, encoding_aes_key, corpid)
            encrypt, msg_signature, timestamp, nonce = crypto.encrypt_message(message_data)
            
            log_data["加密后数据"] = encrypt
            log_data["msg_signature"] = msg_signature
            log_data["timestamp"] = timestamp
            log_data["nonce"] = nonce
            
            # 构建查询参数
            params = {
                "msg_signature": msg_signature,
                "timestamp": timestamp,
                "nonce": nonce
            }
            
            # 构建XML格式的请求体
            xml_data = f"""<xml>
<Encrypt><![CDATA[{encrypt}]]></Encrypt>
<MsgSignature><![CDATA[{msg_signature}]]></MsgSignature>
<TimeStamp>{timestamp}</TimeStamp>
<Nonce><![CDATA[{nonce}]]></Nonce>
</xml>"""
            
            print(f"\n方式1 - 加密发送（XML格式）")
            print(f"查询参数: {params}")
            print(f"请求体: {xml_data}")
            
            log_data["方式1查询参数"] = params
            log_data["方式1请求体(XML)"] = xml_data
            
            response = requests.post(
                callback_url,
                data=xml_data.encode('utf-8'),
                params=params,
                headers={"Content-Type": "application/xml"},
                timeout=10
            )
            
            log_data["方式1响应状态码"] = response.status_code
            log_data["方式1响应头"] = dict(response.headers)
            log_data["方式1响应体"] = response.text
            
            print(f"方式1 - 状态码: {response.status_code}")
            print(f"方式1 - 响应头: {dict(response.headers)}")
            print(f"方式1 - 响应体: {response.text}")
            
            if response.status_code == 200:
                print(f"\n{'='*80}")
                print("邮件数据发送到企业微信接收消息服务器成功（方式1）")
                print(f"响应内容: {response.text}")
                print(f"{'='*80}\n")
                log_data["结果"] = "成功（方式1）"
                save_request_log(log_data)
                return True
        except Exception as e:
            print(f"方式1失败: {str(e)}")
            log_data["方式1异常"] = str(e)
        
        # 尝试方式2：加密数据放在JSON中
        if 'response' not in locals() or response.status_code != 200:
            print(f"\n方式1失败，尝试方式2...")
            try:
                crypto = WeComCrypto(token, encoding_aes_key, corpid)
                encrypt, msg_signature, timestamp, nonce = crypto.encrypt_message(message_data)
                
                params = {
                    "msg_signature": msg_signature,
                    "timestamp": timestamp,
                    "nonce": nonce
                }
                
                data = {
                    "encrypt": encrypt,
                    "msg_signature": msg_signature,
                    "timestamp": timestamp,
                    "nonce": nonce
                }
                
                log_data["方式2查询参数"] = params
                log_data["方式2请求体"] = data
                
                print(f"\n方式2 - 加密发送（JSON格式）")
                print(f"查询参数: {params}")
                print(f"请求体: {data}")
                
                response = requests.post(
                    callback_url,
                    json=data,
                    params=params,
                    headers={"Content-Type": "application/json"},
                    timeout=10
                )
                
                log_data["方式2响应状态码"] = response.status_code
                log_data["方式2响应头"] = dict(response.headers)
                log_data["方式2响应体"] = response.text
                
                print(f"方式2 - 状态码: {response.status_code}")
                print(f"方式2 - 响应头: {dict(response.headers)}")
                print(f"方式2 - 响应体: {response.text}")
                
                if response.status_code == 200:
                    print(f"\n{'='*80}")
                    print("邮件数据发送到企业微信接收消息服务器成功（方式2）")
                    print(f"响应内容: {response.text}")
                    print(f"{'='*80}\n")
                    log_data["结果"] = "成功（方式2）"
                    save_request_log(log_data)
                    return True
            except Exception as e:
                print(f"方式2失败: {str(e)}")
                log_data["方式2异常"] = str(e)
        
        # 尝试方式3：直接发送原始数据（不加密）
        if 'response' not in locals() or response.status_code != 200:
            print(f"\n方式2失败，尝试方式3（不加密）...")
            
            params = {
                "token": token,
                "encoding_aes_key": encoding_aes_key
            }
            
            log_data["方式3查询参数"] = params
            log_data["方式3请求体"] = message_data
            
            print(f"\n方式3 - 不加密发送")
            print(f"查询参数: {params}")
            print(f"请求体: {message_data}")
            
            response = requests.post(
                callback_url,
                json=message_data,
                params=params,
                headers={"Content-Type": "application/json"},
                timeout=10
            )
            
            log_data["方式3响应状态码"] = response.status_code
            log_data["方式3响应头"] = dict(response.headers)
            log_data["方式3响应体"] = response.text
            
            print(f"方式3 - 状态码: {response.status_code}")
            print(f"方式3 - 响应头: {dict(response.headers)}")
            print(f"方式3 - 响应体: {response.text}")
            
            if response.status_code == 200:
                print(f"\n{'='*80}")
                print("邮件数据发送到企业微信接收消息服务器成功（方式3）")
                print(f"响应内容: {response.text}")
                print(f"{'='*80}\n")
                log_data["结果"] = "成功（方式3）"
                save_request_log(log_data)
                return True
        
        # 所有方式都失败
        log_data["最终响应状态码"] = response.status_code
        log_data["最终响应体"] = response.text
        
        print(f"\n{'='*80}")
        print(f"邮件数据发送到企业微信接收消息服务器失败：{response.status_code} - {response.text}")
        print(f"{'='*80}\n")
        log_data["结果"] = "失败"
        save_request_log(log_data)
        return False
        
    except Exception as e:
        print(f"\n{'='*80}")
        print(f"发送到企业微信接收消息服务器时出错：{str(e)}")
        print(f"{'='*80}\n")
        log_data["异常"] = str(e)
        log_data["结果"] = "异常"
        save_request_log(log_data)
        return False
