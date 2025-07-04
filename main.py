import os
import json
import time
import pythoncom
import win32com.client
import datetime

# pip install pywin32

# Load config
with open("config.json", "r") as f:
    config = json.load(f)

SENDER_FILTER = config["sender_email"].lower()
SAVE_FOLDER = config["save_path"]

if not os.path.exists(SAVE_FOLDER):
    os.makedirs(SAVE_FOLDER)

class MailEventHandler:
    def OnItemAdd(self, item):
        print("\n New mail received.")
        try:
            if item.Class != 43:
                print(" Not a MailItem.")
                return

            sender_email = ""

            try:
                sender_email = item.SenderEmailAddress or ""
            except:
                pass

            try:
                if item.Sender.Type == "EX":
                    sender_email = item.Sender.GetExchangeUser().PrimarySmtpAddress
            except:
                pass

            sender_email = sender_email.lower()

            if sender_email == SENDER_FILTER:
                print("Matching sender. Downloading attachments...")
                for attachment in item.Attachments:
                    filename = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{attachment.FileName}"
                    save_path = os.path.join(SAVE_FOLDER, filename)
                    attachment.SaveAsFile(save_path)
                    print(f"Attachment saved: {filename}")
            else:
                print("Mail from non-matching sender. Ignored.")
        except Exception as e:
            print(f"Error processing email: {e}")

def main():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox_event_handlers = []

    print(" Searching all inboxes across all Outlook accounts...\n")

    for i in range(outlook.Folders.Count):
        store = outlook.Folders.Item(i + 1)
        try:
            inbox = store.Folders["Inbox"]
            items = inbox.Items
            #print(f"Monitoring Inbox: {store.Name}")
            handler = win32com.client.WithEvents(items, MailEventHandler)
            inbox_event_handlers.append(handler)
        except Exception as e:
            print(f"Could not access Inbox for {store.Name}: {e}")

    print("\n Now monitoring all inboxes for new emails... Press Ctrl+C to stop.\n")

    while True:
        pythoncom.PumpWaitingMessages()


if __name__ == "__main__":
    main()
