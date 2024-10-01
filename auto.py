import pandas as pd
import win32com.client
import os
import glob

# 讀取 Excel 文件，使用環境變數替代硬編碼路徑
file_path = os.getenv("file_path")
df = pd.read_excel(file_path)

# 讀取草稿模板
def get_draft_template():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    drafts_folder = namespace.GetDefaultFolder(16)
    for item in drafts_folder.Items:
        if item.Subject == "draft template":
            return item
    return None

template_mail = get_draft_template()
template_body = template_mail.HTMLBody

# 使用環境變數替代附件資料夾路徑
attachment_folder = os.getenv("attachment_folder")

# 根據 `no` 欄位的值找到對應的 Excel 文件
def find_attachment(no):
    search_pattern = os.path.join(attachment_folder, f"{no}_*.xls*")
    files = glob.glob(search_pattern)
    if files:
        return files[0]  # 返回找到的第一個文件
    return None

# 生成客製化郵件
def create_custom_drafts(df, template_body):
    outlook = win32com.client.Dispatch("Outlook.Application")
    for index, row in df.iterrows():
        school = row['school']
        name = row['name']
        no = str(row['no'])  # 將 `no` 轉換為字符串
        email = row['email']
        report = row['report'] if pd.notna(row['report']) else ""  # 確保 `report` 欄位存在且不是 NaN
        
        # 替換模板中的佔位符
        custom_body = template_body.replace("{school}", school)
        custom_body = custom_body.replace("{name}", name)
        custom_body = custom_body.replace("{no}", no)
        custom_body = custom_body.replace("{report}", report)
        
        # 創建新的草稿郵件
        mail = outlook.CreateItem(0)
        mail.Subject = os.getenv("MAIL_SUBJECT")  # 使用環境變數替代郵件主旨
        mail.HTMLBody = custom_body
        mail.To = email  # 添加收件者電子郵件地址
        
        # 找到對應的附件並添加
        attachment_path = find_attachment(no)
        if attachment_path:
            mail.Attachments.Add(attachment_path)
        
        mail.Save()
        print(f"Draft email created for {school} {name} with attachment {attachment_path} and sent to {email}")

create_custom_drafts(df, template_body)
