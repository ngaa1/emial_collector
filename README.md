# Outlook邮件读取工具

这是一个简单的Python应用，用于读取本地Outlook中的邮件，支持读取未读邮件和指定时间的邮件。

## 功能特性

- 读取Outlook中的未读邮件
- 按指定时间范围读取邮件（最近几小时或几天）
- 支持读取不同文件夹的邮件（如收件箱、已发送等）
- 将邮件信息保存到文件
- 命令行参数支持，使用灵活

## 环境要求

- Windows操作系统
- Microsoft Outlook已安装并运行
- Python 3.6+

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

1. **读取未读邮件**
   ```bash
   python outlook_reader.py --unread
   ```

2. **读取所有邮件**
   ```bash
   python outlook_reader.py --all
   ```

3. **读取最近24小时的邮件**
   ```bash
   python outlook_reader.py --hours 24 --all
   ```

4. **读取最近3天的邮件**
   ```bash
   python outlook_reader.py --days 3 --all
   ```

5. **读取特定文件夹的邮件**
   ```bash
   python outlook_reader.py --folder "Sent Items" --all
   ```

6. **将结果保存到文件**
   ```bash
   python outlook_reader.py --unread --output emails.txt
   ```

7. **限制最大邮件数量**
   ```bash
   python outlook_reader.py --unread --max 10
   ```

### 命令行参数说明

- `--folder`: 指定Outlook文件夹名称，默认为"Inbox"
- `--unread`: 只读取未读邮件
- `--all`: 读取所有邮件，包括已读
- `--hours`: 读取最近N小时的邮件
- `--days`: 读取最近N天的邮件
- `--max`: 最大邮件数量，默认50
- `--output`: 将结果保存到指定文件

## 打包成可执行文件

为了方便在其他电脑上使用，可以将应用打包成可执行文件。使用PyInstaller：

1. 安装PyInstaller
   ```bash
   pip install pyinstaller
   ```

2. 打包应用
   ```bash
   pyinstaller --onefile outlook_reader.py
   ```

3. 执行文件将生成在`dist`目录中

## 注意事项

- 使用前请确保Outlook已打开并处于运行状态
- 访问其他文件夹时，文件夹名称必须与Outlook界面中的名称完全一致（包括大小写、空格）
- 第一次运行时，Outlook可能会弹出安全提示，需要允许访问
- 打包后的可执行文件需要在安装了Outlook的Windows系统上运行

## 示例输出

```
正在读取Outlook邮件...
文件夹: Inbox
只读取未读: True
时间范围: 2024-01-01 12:00:00 之后
最大邮件数: 50
================================================================================
共找到 5 封邮件
================================================================================
发件人: John Doe
主题: 会议通知
时间: 2024-01-01 14:30:00
内容预览: 尊敬的同事，明天下午3点将召开项目启动会议，请准时参加...
--------------------------------------------------------------------------------
发件人: Jane Smith
主题: 周报
时间: 2024-01-01 13:15:00
内容预览: 本周工作进展如下...
--------------------------------------------------------------------------------
```
