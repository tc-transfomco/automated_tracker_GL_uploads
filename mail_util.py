import win32com.client as win32
import os

from dotenv import load_dotenv

load_dotenv()


# # 1. Spin up Outlook and get your namespace
# olApp = win32.Dispatch("Outlook.Application")
# olNS  = olApp.GetNamespace("MAPI")

# # 2. Find the Account object matching your address
# account = None
# for acct in olNS.Accounts:
#     if acct.SmtpAddress.lower() == os.getenv("email"):
#         account = acct
#         break
# if account is None:
#     raise RuntimeError("Couldn’t find your account in Outlook.Session.Accounts")

# # 3. Create the mail item
# mailItem = olApp.CreateItem(0)  # 0 = olMailItem
# mailItem.Subject    = "Test Email from Python"
# mailItem.BodyFormat =  1        # 1 = olFormatPlain
# mailItem.Body        = "This is a test email sent from Python using win32com."
# mailItem.To          = os.getenv("email_to")

# # 4. Assign the sending account directly
# mailItem.SendUsingAccount = account

# # 5. Send it!
# mailItem.Send()


class MailUtil:
    def __init__(self, to_list="to_list"):
        # 1. Spin up Outlook and get your namespace
        self.olApp = win32.Dispatch("Outlook.Application")
        olNS  = self.olApp.GetNamespace("MAPI")
        
        # 2. Find the Account object matching your address
        self.account = None

        for acct in olNS.Accounts:
            if acct.SmtpAddress.lower() == os.getenv("email"):
                self.account = acct
                break
        if self.account is None:
            raise RuntimeError("Couldn’t find your account in Outlook.Session.Accounts")
        
        self.to_list = os.getenv(to_list).split(',')
    
    def send(self, subject:str, text_body:str, html_body:str = None):
        if len(self.to_list) == 0:
            print("No emails to send to.")
        
        to_field = ";".join(self.to_list)

        mail_item = self.olApp.CreateItem(0)  # olMailItem
        mail_item.Subject    = subject
        mail_item.BodyFormat = 1         # olFormatPlain
        mail_item.Body       = text_body
        if html_body:
            mail_item.HTMLBody = html_body
        mail_item.SendUsingAccount = self.account
        mail_item.To = to_field

        try:
            mail_item.Send()
            print(f"Email sent to: {to_field}\n\n{text_body}")
        except Exception as e:
            print(f"Failed to send to: {to_field}\n{e}")