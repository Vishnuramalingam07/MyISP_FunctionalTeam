import win32com.client
import os
from datetime import datetime

def download_uat_excel_from_outlook(
    subject_keyword="RED: UAT Status for March 28th Release",
    attachment_keyword="UAT Release Detailed Report 28th March",
    save_folder=r"C:\Users\vishnu.ramalingam\MyISP_Tools\Email_Download"
):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox

    print(f"🔍 Searching inbox for: '{subject_keyword}'...")

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # Newest first

    today = datetime.now().date()
    found = 0
    for msg in messages:
        try:
            subject = msg.Subject or ""
            if subject_keyword.lower() not in subject.lower():
                continue

            # Only process emails received today
            received_date = msg.ReceivedTime.date() if hasattr(msg.ReceivedTime, 'date') else datetime.strptime(str(msg.ReceivedTime)[:10], "%Y-%m-%d").date()
            if received_date != today:
                print(f"⏭️  Skipping (not today): {subject} ({msg.ReceivedTime})")
                continue

            print(f"📧 Found email: {subject} ({msg.ReceivedTime})")

            for attachment in msg.Attachments:
                name = attachment.FileName
                if attachment_keyword.lower() in name.lower() and name.endswith(('.xlsx', '.xlsm', '.xls')):
                    save_path = os.path.join(save_folder, name)
                    attachment.SaveAsFile(save_path)
                    print(f"✅ Saved: {save_path}")
                    found += 1

        except Exception as e:
            continue

    if found == 0:
        print("❌ No matching email/attachment found.")
    return found

if __name__ == "__main__":
    download_uat_excel_from_outlook()