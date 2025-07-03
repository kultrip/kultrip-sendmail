import pandas as pd
import argparse
import os
import requests
import base64
from dotenv import load_dotenv
import msal

SENT_EMAILS_FILE = "sent_emails.txt"

def filter_valid_emails(df):
    df = df[df['Email'].notnull()]
    df = df[df['Email'].str.contains('@', na=False)]
    return df

def load_sent_emails(filename):
    if not os.path.exists(filename):
        return set()
    with open(filename, 'r', encoding='utf-8') as f:
        return set(line.strip().lower() for line in f if line.strip())

def save_sent_email(filename, email):
    with open(filename, 'a', encoding='utf-8') as f:
        f.write(email.lower() + "\n")

def get_access_token(client_id, tenant_id, scopes):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id, authority=authority)
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result:
        result = app.acquire_token_interactive(scopes=scopes)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Could not obtain access token: " + str(result))

def send_graph_email(access_token, recipient, subject, html_body, attachment_path=None):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                }
            ]
        },
        "saveToSentItems": "true"
    }
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            content_bytes = base64.b64encode(f.read()).decode('utf-8')
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(attachment_path),
            "contentId": os.path.basename(attachment_path),  # This must match the cid in your HTML
            "isInline": True,
            "contentBytes": content_bytes
        }
        message["message"]["attachments"] = [attachment]

    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    resp = requests.post(url, headers=headers, json=message)
    if resp.status_code >= 400:
        print(f"Error sending email to {recipient}: {resp.status_code} {resp.text}")
        return False
    else:
        print(f"Sent to {recipient}")
        return True

def main():
    # Load environment variables from .env
    load_dotenv()
    client_id = os.getenv('AZURE_APP_CLIENT_ID')
    tenant_id = os.getenv('AZURE_APP_TENANT_ID')
    sender = os.getenv('AZURE_APP_SENDER')  # Your Microsoft 365 email

    parser = argparse.ArgumentParser(description="Bulk mails using Microsoft Graph API")
    parser.add_argument('--max', type=int, default=10, help='Maximum number of mails to be sent (if there are fewer, sends only those)')
    parser.add_argument('--template', type=str, default='email_template.html', help='HTML template')
    parser.add_argument('--start', type=int, default=0, help='Initial index (base 0) in contacts.tsv')
    parser.add_argument('--attachment', type=str, default='signature.png', help='Attachment file path')
    parser.add_argument('--sentlog', type=str, default=SENT_EMAILS_FILE, help='File to store sent emails')
    args = parser.parse_args()

    if not client_id or not tenant_id or not sender:
        raise ValueError("AZURE_APP_CLIENT_ID, AZURE_APP_TENANT_ID, and AZURE_APP_SENDER must be set in the .env file.")

    sent_emails = load_sent_emails(args.sentlog)

    contacts = pd.read_csv('contacts.tsv', sep='\t')
    contacts = filter_valid_emails(contacts)
    contacts = contacts.iloc[args.start:args.start+args.max]

    with open(args.template, 'r', encoding='utf-8') as f:
        html_template = f.read()

    scopes = ["Mail.Send"]
    access_token = get_access_token(client_id, tenant_id, scopes)

    for idx, row in contacts.iterrows():
        recipient_email = row['Email'].strip().lower()
        if recipient_email in sent_emails:
            print(f"Skipping {recipient_email} (already sent)")
            continue
        html_body = html_template
        for col in contacts.columns:
            html_body = html_body.replace(f"{{{{{col}}}}}", str(row[col]))
        subject = f"{row['Name']}, crie roteiros personalizados em minutos para seus clientes"
        if send_graph_email(
            access_token=access_token,
            recipient=recipient_email,
            subject=subject,
            html_body=html_body,
            attachment_path=args.attachment
        ):
            save_sent_email(args.sentlog, recipient_email)
            sent_emails.add(recipient_email)

    print("All messages sent.")

if __name__ == '__main__':
    main()