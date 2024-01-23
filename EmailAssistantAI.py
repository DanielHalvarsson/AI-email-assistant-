import win32com.client
import os
from openai import OpenAI
from dotenv import load_dotenv
load_dotenv()

# Initialize OpenAI client
client = OpenAI()

def get_smtp_address(message):
    try:
        # Check if the sender is an Exchange user
        if message.SenderEmailType == "EX":
            sender = message.Sender.GetExchangeUser()
            if sender:
                return sender.PrimarySmtpAddress
        else:
            return message.SenderEmailAddress
    except Exception as e:
        print(f"Error getting SMTP address: {e}")
        return None


def get_latest_email_content():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        if len(messages) > 0:
            latest_message = messages[0]
            recipient = get_smtp_address(latest_message)
            return latest_message.Subject, latest_message.Body, recipient
        else:
            print("No emails found in the inbox.")
            return None, None, None
    except Exception as e:
        print(f"Error retrieving the latest email: {e}")
        return None, None, None

def generate_ai_response(email_subject, email_body):
    try:
        # Modify this part to use the email subject and body as part of the input
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo-1106",
            messages=[
                {"role": "system", "content": "You are a helpfull email-assistant that reads emails to prepare a draft responds in the same language as in the original email."},
                {"role": "user", "content": email_body},
            ]
        )
        return completion.choices[0].message.content
    except Exception as e:
        print(f"Error generating AI response: {e}")
        return ""

def create_draft_email(response, recipient, email_subject):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        draft = outlook.CreateItem(0)  # 0 is the code for a mail item
        draft.To = recipient
        draft.Subject = "RE: " + email_subject
        draft.Body = response
        draft.Save()
        print(f"Draft email to {recipient} created successfully.")
    except Exception as e:
        print(f"Error creating draft email: {e}")

# Main logic
email_subject, email_body, recipient = get_latest_email_content()
ai_response = generate_ai_response(email_subject, email_body)
create_draft_email(ai_response, recipient, email_subject)