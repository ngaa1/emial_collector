import win32com.client
import win32timezone
from datetime import datetime, timedelta
import argparse
import os
import json
import schedule
import time
import threading
import requests

def get_outlook_folders():
    """
    获取Outlook中的所有邮件文件夹
    :return: 文件夹名称列表
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        
        # 获取默认邮箱的所有文件夹
        folders = []
        for folder in namespace.Folders(inbox.Name).Folders:
            folders.append(folder.Name)
        
        return folders
    except Exception as e:
        print(f"获取文件夹列表时出错：{str(e)}")
        return ["Inbox"]

def read_outlook_emails(
    folder_name="Inbox",
    read_unread_only=True,            # 是否只读取未读邮件
    since_datetime=None,              # 只读取指定时间之后的邮件
    max_emails=50                     # 一次最多读取多少封
):
    """
    读取Outlook邮件
    :param folder_name: 收件箱文件夹名称（如 "Inbox", "Sent Items" 等）
    :param read_unread_only: 是否只读取未读邮件
    :param since_datetime: datetime类型，只读取这个时间之后的邮件
    :param max_emails: 最大邮件数
    :return: 邮件列表
    """
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # 获取收件箱
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        if folder_name != "Inbox":
            # 如果要访问其他文件夹（如已发送、草稿等），需要用名称精确匹配
            try:
                inbox = namespace.Folders(inbox.Name).Folders(folder_name)
            except Exception as e:
                print(f"错误：找不到文件夹 '{folder_name}'，请检查文件夹名称是否正确")
                return []

        messages = inbox.Items
        # 按日期倒序排列（最新的在前）
        messages.Sort("[ReceivedTime]", True)

        # 筛选逻辑
        filtered_messages = []
        for msg in messages:
            # 根据"未读"状态筛选
            if read_unread_only and msg.UnRead:
                pass  # 未读邮件，保留
            elif not read_unread_only:
                pass  # 所有邮件都保留
            else:
                continue  # 未读标志为False且要求未读时，跳过

            # 根据"发送时间"筛选
            if since_datetime:
                # 简单处理时间比较，避免时区问题
                try:
                    # 尝试获取ReceivedTime属性
                    received_time = msg.ReceivedTime
                    # 尝试直接比较
                    if received_time < since_datetime:
                        continue
                except Exception:
                    # 如果获取时间失败，跳过时间筛选
                    pass

            filtered_messages.append(msg)

            if len(filtered_messages) >= max_emails:
                break

        return filtered_messages
    except Exception as e:
        print(f"错误：{str(e)}")
        print("提示：请确保Outlook已打开并运行")
        return []

def print_email_info(msg):
    """打印邮件信息"""
    try:
        sender = msg.SenderName
        subject = msg.Subject
        received_time = msg.ReceivedTime
        body = msg.Body[:200] + "..." if len(msg.Body) > 200 else msg.Body

        print(f"发件人: {sender}")
        print(f"主题: {subject}")
        print(f"时间: {received_time}")
        print(f"内容预览: {body}")
        print("-" * 80)
    except Exception as e:
        print(f"打印邮件信息时出错：{str(e)}")
        print("-" * 80)

def save_emails_to_file(emails, output_file):
    """将邮件信息保存到文件"""
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"邮件总数: {len(emails)}\n")
            f.write(f"生成时间: {datetime.now()}\n")
            f.write("=" * 80 + "\n")
            
            for i, msg in enumerate(emails, 1):
                f.write(f"邮件 {i}:\n")
                f.write(f"发件人: {msg.SenderName}\n")
                f.write(f"主题: {msg.Subject}\n")
                f.write(f"时间: {msg.ReceivedTime}\n")
                f.write(f"内容: {msg.Body}\n")
                f.write("=" * 80 + "\n")
        print(f"邮件信息已保存到: {output_file}")
    except Exception as e:
        print(f"保存文件时出错：{str(e)}")

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
        
        # 发送消息
        url = f"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={access_token}"
        data = {
            "touser": touser,
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

def get_user_choice(prompt, options, default=None):
    """获取用户选择"""
    print(prompt)
    for i, option in enumerate(options, 1):
        print(f"{i}. {option}")
    while True:
        default_text = f"，默认 {default}" if default else ""
        choice = input(f"请选择 (1-{len(options)}){default_text}: ").strip()
        if not choice and default:
            return default
        try:
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(options):
                return options[choice_idx]
            print(f"请输入有效的选项 (1-{len(options)})")
        except ValueError:
            print("请输入数字")

def get_config_file():
    """获取配置文件路径"""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def get_wechat_config_file():
    """获取企业微信配置文件路径"""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "wechat_config.json")

def load_wechat_config():
    """从文件加载企业微信配置"""
    try:
        config_file = get_wechat_config_file()
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config
    except Exception as e:
        print(f"加载企业微信配置时出错：{str(e)}")
    return {}

def get_head_content():
    """获取head.txt文件的内容"""
    try:
        head_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "head.txt")
        if os.path.exists(head_file):
            with open(head_file, 'r', encoding='utf-8') as f:
                return f.read().strip()
    except Exception as e:
        print(f"读取head.txt时出错：{str(e)}")
    return ""

def save_config(config):
    """保存配置到文件"""
    try:
        config_file = get_config_file()
        with open(config_file, 'w', encoding='utf-8') as f:
            # 转换datetime对象为字符串
            config_copy = config.copy()
            if config_copy.get('since_datetime'):
                config_copy['since_datetime'] = str(config_copy['since_datetime'])
            json.dump(config_copy, f, ensure_ascii=False, indent=2)
        print(f"配置已保存到: {config_file}")
    except Exception as e:
        print(f"保存配置时出错：{str(e)}")

def load_config():
    """从文件加载配置"""
    try:
        config_file = get_config_file()
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 转换字符串为datetime对象
                if config.get('since_datetime'):
                    try:
                        config['since_datetime'] = datetime.fromisoformat(config['since_datetime'])
                    except:
                        config['since_datetime'] = None
                return config
    except Exception as e:
        print(f"加载配置时出错：{str(e)}")
    return {}

def get_user_input(prompt, default=None, is_int=False):
    """获取用户输入"""
    while True:
        default_text = f"，默认 {default}" if default else ""
        value = input(f"{prompt}{default_text}: ").strip()
        if not value and default:
            return default
        if is_int:
            try:
                return int(value)
            except ValueError:
                print("请输入数字")
        else:
            return value

def interactive_mode():
    """交互式模式"""
    print("================================")
    print("Outlook邮件读取工具 - 交互式模式")
    print("================================")
    
    # 询问是否使用旧设置
    use_saved_config = get_user_choice(
        "是否使用上次的设置:",
        ["是", "否"],
        "是"
    )
    
    if use_saved_config == "是":
        # 加载并返回上次的配置
        saved_config = load_config()
        print("使用上次的设置:")
        print(f"  文件夹: {saved_config.get('folder', 'Inbox')}")
        print(f"  读取模式: {'只读取未读邮件' if saved_config.get('read_unread_only', True) else '读取所有邮件'}")
        if saved_config.get('since_datetime'):
            print(f"  时间范围: {saved_config.get('since_datetime')} 之后")
        print(f"  最大邮件数: {saved_config.get('max_emails', 50)}")
        print(f"  保存到文件: {'是' if saved_config.get('output') else '否'}")
        print(f"  发送到企业微信: {'是' if saved_config.get('wechat_enabled', False) else '否'}")
        print(f"  定时运行: {'是' if saved_config.get('schedule_enabled', False) else '否'}")
        if saved_config.get('schedule_enabled', False):
            print(f"  运行时间: {saved_config.get('schedule_time', '09:00')}")
        print("================================")
        return saved_config
    else:
        # 加载默认配置
        saved_config = {}
        
        # 选择读取模式
        read_mode = get_user_choice(
            "请选择读取模式:",
            ["只读取未读邮件", "读取所有邮件"],
            "只读取未读邮件"
        )
        read_unread_only = read_mode == "只读取未读邮件"
        
        # 选择文件夹
        folder_option = get_user_choice(
            "请选择文件夹选择方式:",
            ["从Outlook中自动扫描", "手动输入文件夹名称"],
            "从Outlook中自动扫描"
        )
        
        if folder_option == "从Outlook中自动扫描":
            print("正在扫描Outlook文件夹...")
            folders = get_outlook_folders()
            if len(folders) > 1:
                folder = get_user_choice(
                    "请选择要读取的文件夹:",
                    folders,
                    "Inbox"
                )
            else:
                folder = folders[0]
                print(f"只找到一个文件夹: {folder}")
        else:
            folder = get_user_input(
                "请输入要读取的文件夹名称",
                "Inbox"
            )
        
        # 选择时间范围
        time_option = get_user_choice(
            "请选择时间范围:",
            ["不限制时间", "最近几小时", "最近几天"],
            "不限制时间"
        )
        
        since_datetime = None
        if time_option == "最近几小时":
            hours = get_user_input("请输入小时数", 24, is_int=True)
            since_datetime = datetime.now() - timedelta(hours=hours)
        elif time_option == "最近几天":
            days = get_user_input("请输入天数", 7, is_int=True)
            since_datetime = datetime.now() - timedelta(days=days)
        
        # 最大邮件数量
        max_emails = get_user_input("请输入最大邮件数量", 50, is_int=True)
        
        # 是否保存到文件
        save_option = get_user_choice(
            "是否保存结果到文件:",
            ["是", "否"],
            "否"
        )
        output_file = None
        if save_option == "是":
            default_file = f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            output_file = get_user_input("请输入保存文件名", default_file)
        
        # 是否发送到企业微信
        wechat_option = get_user_choice(
            "是否发送结果到企业微信:",
            ["是", "否"],
            "否"
        )
        wechat_enabled = wechat_option == "是"
        
        # 企业微信相关配置
        corpid = None
        corpsecret = None
        agentid = None
        touser = None
        
        if wechat_enabled:
            corpid = get_user_input("请输入企业微信企业ID")
            corpsecret = get_user_input("请输入应用Secret")
            agentid = get_user_input("请输入应用ID", is_int=True)
            touser = get_user_input("请输入接收人（userid列表，用|分隔，如：zhangsan|lisi）")
        
        # 构建配置
        config = {
            "folder": folder,
            "read_unread_only": read_unread_only,
            "since_datetime": since_datetime,
            "max_emails": max_emails,
            "output": output_file,
            "schedule_enabled": False,
            "schedule_time": "09:00",
            "wechat_enabled": wechat_enabled,
            "corpid": corpid,
            "corpsecret": corpsecret,
            "agentid": agentid,
            "touser": touser
        }
        
        # 保存配置
        save_config(config)
        
        return config

def run_email_reader(config):
    """运行邮件读取器"""
    print(f"\n[定时任务] 开始读取Outlook邮件...")
    print(f"文件夹: {config['folder']}")
    print(f"只读取未读: {config['read_unread_only']}")
    if config.get('since_datetime'):
        print(f"时间范围: {config['since_datetime']} 之后")
    print(f"最大邮件数: {config['max_emails']}")
    print("=" * 80)
    
    emails = read_outlook_emails(
        folder_name=config['folder'],
        read_unread_only=config['read_unread_only'],
        since_datetime=config.get('since_datetime'),
        max_emails=config['max_emails']
    )
    
    # 显示邮件信息
    print(f"共找到 {len(emails)} 封邮件")
    print("=" * 80)
    
    for msg in emails:
        print_email_info(msg)
    
    # 保存到文件
    if config.get('output'):
        output_file = f"emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        save_emails_to_file(emails, output_file)
    
    # 发送到企业微信
    if config.get('wechat_enabled'):
        # 优先使用独立配置文件中的参数
        wechat_config = load_wechat_config()
        corpid = wechat_config.get('corpid') or config.get('corpid')
        corpsecret = wechat_config.get('corpsecret') or config.get('corpsecret')
        agentid = wechat_config.get('agentid') or config.get('agentid')
        touser = wechat_config.get('touser') or config.get('touser')
        if corpid and corpsecret and agentid and touser:
            send_wechat_message(corpid, corpsecret, agentid, touser, emails)
        else:
            print("企业微信配置不完整，无法发送消息")
    
    print("[定时任务] 操作完成！")

def start_scheduler(config, schedule_time):
    """启动定时任务调度器"""
    def job():
        run_email_reader(config)
    
    # 设置每天的定时任务
    schedule.every().day.at(schedule_time).do(job)
    
    print(f"\n定时任务已设置：每天 {schedule_time} 运行")
    print("按 Ctrl+C 退出...")
    
    # 立即运行一次
    run_email_reader(config)
    
    # 循环执行定时任务
    while True:
        schedule.run_pending()
        time.sleep(60)

def main():
    import sys
    
    # 检查是否有命令行参数
    if len(sys.argv) > 1:
        # 使用命令行参数模式
        parser = argparse.ArgumentParser(description="Outlook邮件读取工具")
        parser.add_argument("--folder", default="Inbox", help="指定文件夹名称，默认为Inbox")
        parser.add_argument("--unread", action="store_true", help="只读取未读邮件")
        parser.add_argument("--all", action="store_true", help="读取所有邮件，包括已读")
        parser.add_argument("--hours", type=int, help="读取最近N小时的邮件")
        parser.add_argument("--days", type=int, help="读取最近N天的邮件")
        parser.add_argument("--max", type=int, default=50, help="最大邮件数量，默认50")
        parser.add_argument("--output", help="将结果保存到文件")
        parser.add_argument("--wechat", action="store_true", help="发送结果到企业微信")
        parser.add_argument("--corpid", help="企业微信企业ID")
        parser.add_argument("--corpsecret", help="企业微信应用Secret")
        parser.add_argument("--agentid", type=int, help="企业微信应用ID")
        parser.add_argument("--touser", help="企业微信接收人（userid列表，用|分隔）")
        
        args = parser.parse_args()
        
        # 确定是否只读取未读邮件
        read_unread_only = args.unread
        if args.all:
            read_unread_only = False
        
        # 确定时间筛选条件
        since_datetime = None
        if args.hours:
            since_datetime = datetime.now() - timedelta(hours=args.hours)
        elif args.days:
            since_datetime = datetime.now() - timedelta(days=args.days)
        
        config = {
            "folder": args.folder,
            "read_unread_only": read_unread_only,
            "since_datetime": since_datetime,
            "max_emails": args.max,
            "output": args.output,
            "wechat_enabled": args.wechat,
            "corpid": args.corpid,
            "corpsecret": args.corpsecret,
            "agentid": args.agentid,
            "touser": args.touser
        }
    else:
        # 使用交互式模式
        config = interactive_mode()
    
    # 加载上次的定时设置
    saved_config = load_config()
    schedule_time = saved_config.get('schedule_time', '09:00')
    
    # 询问是否设置定时运行
    schedule_option = get_user_choice(
        "是否设置定时运行:",
        ["是", "否"],
        "是" if saved_config.get('schedule_enabled', False) else "否"
    )
    
    if schedule_option == "是":
        # 输入定时运行时间
        schedule_time = get_user_input("请输入每天运行时间（格式：HH:MM，例如 09:00）", schedule_time)
        # 保存定时设置
        config['schedule_enabled'] = True
        config['schedule_time'] = schedule_time
        save_config(config)
        # 启动定时任务
        start_scheduler(config, schedule_time)
    else:
        # 保存定时设置
        config['schedule_enabled'] = False
        config['schedule_time'] = schedule_time
        save_config(config)
        # 读取邮件
        print(f"正在读取Outlook邮件...")
        print(f"文件夹: {config['folder']}")
        print(f"只读取未读: {config['read_unread_only']}")
        if config.get('since_datetime'):
            print(f"时间范围: {config['since_datetime']} 之后")
        print(f"最大邮件数: {config['max_emails']}")
        print("=" * 80)
        
        emails = read_outlook_emails(
            folder_name=config['folder'],
            read_unread_only=config['read_unread_only'],
            since_datetime=config.get('since_datetime'),
            max_emails=config['max_emails']
        )
        
        # 显示邮件信息
        print(f"共找到 {len(emails)} 封邮件")
        print("=" * 80)
        
        for msg in emails:
            print_email_info(msg)
        
        # 保存到文件
        if config.get('output'):
            save_emails_to_file(emails, config['output'])
        
        # 发送到企业微信
        if config.get('wechat_enabled'):
            # 优先使用独立配置文件中的参数
            wechat_config = load_wechat_config()
            corpid = wechat_config.get('corpid') or config.get('corpid')
            corpsecret = wechat_config.get('corpsecret') or config.get('corpsecret')
            agentid = wechat_config.get('agentid') or config.get('agentid')
            touser = wechat_config.get('touser') or config.get('touser')
            if corpid and corpsecret and agentid and touser:
                send_wechat_message(corpid, corpsecret, agentid, touser, emails)
            else:
                print("企业微信配置不完整，无法发送消息")
        
        # 程序结束时暂停
        print("\n操作完成！")
        input("请按回车键退出...")

if __name__ == "__main__":
    main()
