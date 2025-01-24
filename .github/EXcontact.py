import win32com.client
import openpyxl
import datetime
import pytz
import os

# Define the timezone (example: use UTC)
timezone = pytz.timezone('UTC')

# Define the time range with timezone information
start_date = timezone.localize(datetime.datetime(2024, 7, 1))
end_date = timezone.localize(datetime.datetime(2024, 7, 4))

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
sent_items = namespace.GetDefaultFolder(5)  # 5 corresponds to the Sent Items folder

# Prepare Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Recipients"
ws.append(["Subject", "Recipient Name", "Email Address", "Sent On"])

# Function to resolve email address
def get_smtp_address(recipient):
    try:
        address_entry = recipient.AddressEntry
        if address_entry.Type == "EX":
            exchange_user = address_entry.GetExchangeUser()
            if exchange_user:
                return exchange_user.PrimarySmtpAddress
        return recipient.Address
    except Exception as e:
        print(f"Could not resolve email address for {recipient.Name}: {e}")
        return recipient.Address

# Iterate through the emails and extract details
for item in sent_items.Items:
    if item.Class == 43:  # 43 corresponds to MailItem
        sent_on = item.SentOn

        # Ensure sent_on is offset-aware
        if sent_on.tzinfo is None:
            sent_on = timezone.localize(sent_on)

        if start_date <= sent_on <= end_date:
            subject = item.Subject
            for recipient in item.Recipients:
                recipient_name = recipient.Name
                email_address = get_smtp_address(recipient)

                # Convert sent_on to naive datetime
                sent_on_naive = sent_on.replace(tzinfo=None)

                ws.append([subject, recipient_name, email_address, sent_on_naive])

# Define the file path
excel_path = "D:\\workroom\\recipients.xlsx"

# Ensure the directory exists
os.makedirs(os.path.dirname(excel_path), exist_ok=True)

# Save the workbook
wb.save(excel_path)

print(f"Export complete. Excel file saved to {excel_path}")