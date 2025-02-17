import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber  # 更强大的PDF文本提取库
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import ssl

# 邮箱配置（根据中科院邮箱要求修改）
SMTP_SERVER = "mail.cstnet.cn"  # 根据实际服务器修改
SMTP_PORT = 465  # SSL端口
EMAIL_ADDRESS = "zhang_cheng@mail.sim.ac.cn"  # 发送邮箱
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "default_password")  # 邮箱客户端专用密码/授权码,使用环境变量获取
RECIPIENT_EMAIL = "zhang_cheng@mail.sim.ac.cn"  # 收件人邮箱

# ==============================================

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        process_pdfs(folder_path)

def extract_invoice_info(text):
    """使用正则表达式提取发票信息"""
    patterns = {
        'invoice_number': [
            r'发票号码[:：]\s*([A-Z0-9]{8,20})',
            r'号码[:：]\s*(\d{8,20})',
            r'No\.([A-Z0-9]{8,20})'
        ],
        'total_amount': [
            r'价税合计[^\d]*([\d,]+\.\d{2})',
            r'合计[^\d]*([\d,]+\.\d{2})',
            r'Amount[\s:：]*([\d,]+\.\d{2})'
        ]
    }

    invoice_number = None
    total_amount = None

    # 尝试多种匹配模式
    for pattern in patterns['invoice_number']:
        match = re.search(pattern, text)
        if match:
            invoice_number = match.group(1).strip()
            break

    for pattern in patterns['total_amount']:
        match = re.search(pattern, text)
        if match:
            total_amount = match.group(1).replace(',', '')
            break

    return invoice_number, total_amount


def print_pdf(pdf_path):
    """使用Selenium打开PDF文件并发送打印命令"""
    try:
        # 设置 Chrome 的选项
        options = Options()
        options.add_argument("--kiosk-printing")  # 开启无提示打印模式
        # options.add_argument('--headless')  # 如果不想显示浏览器窗口，可以取消注释
        # 设置ChromeDriver的路径
        service = Service('C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')  # 替换为你的ChromeDriver路径

        #强烈建议关闭chrome的自动更新，不然会一直烦你同步更新chromedriver

        # 创建 WebDriver 实例，指定驱动程序路径
        with webdriver.Chrome(service=service,options=options) as driver:
            # 打开 PDF 文件
            driver.get(f"file:///{pdf_path}")

            # 等待页面加载完成
            time.sleep(2)

            # 执行打印命令
            driver.execute_script('window.print();')

            # 等待打印完成（根据实际情况调整等待时间）
            time.sleep(3)

    except Exception as e:
        print(f"打印错误: {str(e)}")
    finally:
        # 确保关闭浏览器
        driver.quit()

def send_email(folder_path):
    """发送folder_path下所有PDF文件作为附件"""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = f"张程-{datetime.now().strftime('%Y-%m-%d')}-发票"

    # 添加所有PDF附件
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        raise ValueError("没有可发送的PDF文件")

    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)

    try:
        # 使用SSL安全连接
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, msg.as_string())
        print(f"邮件发送成功，附件数量：{len(pdf_files)}")
    except Exception as e:
        messagebox.showerror("邮件错误", f"邮件发送失败: {str(e)}")
        raise  # 将异常传递到上层处理

def process_pdfs(folder_path):
    processed_files = []
    errors = []

    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)

            try:
                # 提取文本内容
                text = ""
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        text += page.extract_text()

                # 提取信息
                invoice_number, total_amount = extract_invoice_info(text)
                if not invoice_number or not total_amount:
                    raise ValueError("未能提取完整发票信息")

                # 重命名文件
                new_filename = f"张程+{invoice_number}+{total_amount}元.pdf"
                new_path = os.path.join(folder_path, new_filename)

                # 处理文件名冲突
                counter = 1
                while os.path.exists(new_path):
                    new_filename = f"张程+{invoice_number}+{total_amount}元({counter}).pdf"
                    new_path = os.path.join(folder_path, new_filename)
                    counter += 1

                os.rename(pdf_path, new_path)

                print(new_path)
                # 打印文件
                print_pdf(new_path)

                processed_files.append(new_filename)

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

    # 所有文件处理完成后发送邮件
    try:
        send_email(folder_path)
        processed_files.append("邮件发送成功")
    except Exception as e:
        errors.append(f"邮件发送失败: {str(e)}")

    # 显示处理结果
    result_msg = []
    if processed_files:
        result_msg.append("成功处理文件：\n" + "\n".join(processed_files))
    if errors:
        result_msg.append("\n\n处理失败文件：\n" + "\n".join(errors))

    messagebox.showinfo("处理结果", "\n".join(result_msg) if result_msg else "没有找到PDF文件")

# 创建UI界面
root = tk.Tk()
root.title("发票处理系统 v1.0")
root.geometry("400x200")

frame = tk.Frame(root)
frame.pack(pady=40)

label = tk.Label(frame, text="请选择包含发票PDF的文件夹：")
label.pack()

select_btn = tk.Button(frame, text="选择文件夹", command=select_folder,
                      bg="#4CAF50", fg="white", padx=20, pady=10)
select_btn.pack(pady=10)

root.mainloop()
