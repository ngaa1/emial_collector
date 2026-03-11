import win32com.client
import win32timezone
from datetime import datetime, timedelta
import argparse
import os

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
                # msg.ReceivedTime 是 datetime 对象
                if msg.ReceivedTime < since_datetime:
                    continue

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

def main():
    parser = argparse.ArgumentParser(description="Outlook邮件读取工具")
    parser.add_argument("--folder", default="Inbox", help="指定文件夹名称，默认为Inbox")
    parser.add_argument("--unread", action="store_true", help="只读取未读邮件")
    parser.add_argument("--all", action="store_true", help="读取所有邮件，包括已读")
    parser.add_argument("--hours", type=int, help="读取最近N小时的邮件")
    parser.add_argument("--days", type=int, help="读取最近N天的邮件")
    parser.add_argument("--max", type=int, default=50, help="最大邮件数量，默认50")
    parser.add_argument("--output", help="将结果保存到文件")
    
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
    
    # 读取邮件
    print(f"正在读取Outlook邮件...")
    print(f"文件夹: {args.folder}")
    print(f"只读取未读: {read_unread_only}")
    if since_datetime:
        print(f"时间范围: {since_datetime} 之后")
    print(f"最大邮件数: {args.max}")
    print("=" * 80)
    
    emails = read_outlook_emails(
        folder_name=args.folder,
        read_unread_only=read_unread_only,
        since_datetime=since_datetime,
        max_emails=args.max
    )
    
    # 显示邮件信息
    print(f"共找到 {len(emails)} 封邮件")
    print("=" * 80)
    
    for msg in emails:
        print_email_info(msg)
    
    # 保存到文件
    if args.output:
        save_emails_to_file(emails, args.output)

if __name__ == "__main__":
    main()
