#!/usr/bin/env python3
"""
测试脚本：模拟Outlook邮件读取功能
"""

import sys
import os
from datetime import datetime, timedelta

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

class MockOutlook:
    """模拟Outlook对象"""
    def Dispatch(self, app_name):
        return MockApplication()

class MockApplication:
    """模拟Outlook应用"""
    def GetNamespace(self, namespace):
        return MockNamespace()

class MockNamespace:
    """模拟Outlook命名空间"""
    def GetDefaultFolder(self, folder_id):
        return MockFolder("Inbox")
    
    def Folders(self, folder_name):
        return MockFolders()

class MockFolders:
    """模拟文件夹集合"""
    def Folders(self, folder_name):
        return MockFolder(folder_name)

class MockFolder:
    """模拟文件夹"""
    def __init__(self, name):
        self.name = name
        self.Items = MockItems()

class MockItems:
    """模拟邮件项集合"""
    def __init__(self):
        # 创建模拟邮件
        self.emails = []
        now = datetime.now()
        
        # 未读邮件
        for i in range(3):
            email = MockEmail(
                sender_name=f"发件人{i+1}",
                subject=f"测试邮件{i+1}",
                body=f"这是测试邮件{i+1}的内容",
                received_time=now - timedelta(hours=i),
                unread=True
            )
            self.emails.append(email)
        
        # 已读邮件
        for i in range(2):
            email = MockEmail(
                sender_name=f"发件人{i+4}",
                subject=f"已读邮件{i+1}",
                body=f"这是已读邮件{i+1}的内容",
                received_time=now - timedelta(days=i+1),
                unread=False
            )
            self.emails.append(email)
    
    def Sort(self, field, descending):
        # 模拟排序
        pass
    
    def __iter__(self):
        return iter(self.emails)

class MockEmail:
    """模拟邮件"""
    def __init__(self, sender_name, subject, body, received_time, unread):
        self.SenderName = sender_name
        self.Subject = subject
        self.Body = body
        self.ReceivedTime = received_time
        self.UnRead = unread

def test_read_outlook_emails():
    """测试读取邮件功能"""
    print("开始测试Outlook邮件读取功能...")
    
    # 模拟win32com.client
    import win32com.client
    original_dispatch = win32com.client.Dispatch
    win32com.client.Dispatch = MockOutlook().Dispatch
    
    try:
        from outlook_reader import read_outlook_emails, print_email_info
        
        # 测试1: 读取未读邮件
        print("\n测试1: 读取未读邮件")
        unread_emails = read_outlook_emails(read_unread_only=True, max_emails=10)
        print(f"找到 {len(unread_emails)} 封未读邮件")
        for email in unread_emails:
            print_email_info(email)
        
        # 测试2: 读取所有邮件
        print("\n测试2: 读取所有邮件")
        all_emails = read_outlook_emails(read_unread_only=False, max_emails=10)
        print(f"找到 {len(all_emails)} 封邮件")
        for email in all_emails:
            print_email_info(email)
        
        # 测试3: 读取最近1小时的邮件
        print("\n测试3: 读取最近1小时的邮件")
        one_hour_ago = datetime.now() - timedelta(hours=1)
        recent_emails = read_outlook_emails(
            since_datetime=one_hour_ago,
            read_unread_only=False,
            max_emails=10
        )
        print(f"找到 {len(recent_emails)} 封最近1小时的邮件")
        for email in recent_emails:
            print_email_info(email)
        
        print("\n测试完成！")
        
    finally:
        # 恢复原始函数
        win32com.client.Dispatch = original_dispatch

if __name__ == "__main__":
    test_read_outlook_emails()
