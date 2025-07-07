import time
import os
import logging
import pandas as pd
import win32com.client
from datetime import datetime

# Setup logging
logging.basicConfig(filename='xpor_auto_email.log',
                    level=logging.INFO,
                    format='%(asctime)s %(levelname)s:%(message)s')

# Email settings
RECIPIENT = 'XXX@YYY.ZZZ'
SUBJECT_PREFIX = 'XPOR'
CHECK_INTERVAL = 300  # seconds (5 minutes)
ATTACHMENT_DIR = 'xpor_attachments'

if not os.path.exists(ATTACHMENT_DIR):
    os.makedirs(ATTACHMENT_DIR)

def process_excel(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        if df.shape[0] < 1:
            logging.warning(f"No data rows in {file_path}")
            return None
        # Columns: A=0, B=1, C=2, D=3, E=4
        col_D = df.iloc[:, 3].dropna()
        col_E = df.iloc[:, 4].dropna()
        col_B = df.iloc[:, 1].dropna()
        if col_D.empty or col_E.empty or col_B.empty:
            logging.warning(f"Missing data in columns in {file_path}")
            return None
        median_D = col_D.median()
        bottom_E = col_E.iloc[-1]
        bottom_B = col_B.iloc[-1]
        result1 = median_D * (1 + bottom_E)
        if result1 == 0:
            logging.warning(f"Result1 is zero in {file_path}")
            return None
        final_number = (bottom_B * 100) / result1
        return final_number
    except Exception as e:
        logging.error(f"Error processing {file_path}: {e}")
        return None

def send_email(result, original_subject):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = RECIPIENT
        mail.Subject = f"XPOR Result for {datetime.now().strftime('%Y-%m-%d')}"
        mail.Body = f"The computed XPOR value from today's file (original subject: {original_subject}) is: {result:.4f}"
        mail.Send()
        logging.info(f"Sent result email: {result:.4f}")
    except Exception as e:
        logging.error(f"Error sending email: {e}")

def check_emails():
    try:
        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items
        messages = messages.Restrict(f"[Unread]=true AND [Subject] Like '{SUBJECT_PREFIX}%' AND [Attachments].Count > 0")
        for message in list(messages):
            subject = message.Subject
            received_time = message.ReceivedTime
            logging.info(f"Processing email: {subject} at {received_time}")
            for att in message.Attachments:
                if att.FileName.lower().endswith('.xlsx'):
                    save_path = os.path.join(ATTACHMENT_DIR, f"{received_time.strftime('%Y%m%d_%H%M%S')}_{att.FileName}")
                    att.SaveAsFile(save_path)
                    logging.info(f"Saved attachment to {save_path}")
                    result = process_excel(save_path)
                    if result is not None:
                        send_email(result, subject)
                    else:
                        logging.warning(f"No result computed for {save_path}")
            message.Unread = False  # Mark as read
            message.Save()
            logging.info(f"Marked email as read: {subject}")
    except Exception as e:
        logging.error(f"Error checking emails: {e}")

def main():
    logging.info("XPOR auto email script started.")
    while True:
        check_emails()
        time.sleep(CHECK_INTERVAL)

if __name__ == '__main__':
    main()
