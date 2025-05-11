import os
import pickle
import base64
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import requests
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import html
import re

import markdown
from xhtml2pdf import pisa



import dropbox
import csv

DROPBOX_ACCESS_TOKEN = "your-dropbox-token" #your dropbox access token to store emails
DROPBOX_BASE_PATH = "/gmail_emails"

dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)

def upload_to_dropbox(local_path, dropbox_path):
    with open(local_path, "rb") as f:
        dbx.files_upload(f.read(), dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
    shared_link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_path)
    # Make it a direct link
    return shared_link_metadata.url.replace("?dl=0", "?raw=1")



# CONFIG
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

NUM_EMAILS = 20 #you can set your limit
ATTACHMENT_DIR = 'attachments'  # Folder to save attachments

# ------------------------------------------
# AUTHENTICATE WITH GMAIL
# ------------------------------------------
def authenticate_gmail():
    creds = None
    if os.path.exists('token.pkl'):
        with open('token.pkl', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pkl', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# ------------------------------------------
# FETCH EMAILS + ATTACHMENTS
# ------------------------------------------
def get_message_body(payload):
    if 'parts' in payload:
        for part in payload['parts']:
            if part['mimeType'] == 'text/plain':
                return base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
            elif part['mimeType'].startswith('multipart'):
                return get_message_body(part)
    elif payload['mimeType'] == 'text/html':
        return base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8')
    return '(No plain text body found)'

def sanitize_folder_name(name):
    """Remove invalid folder name characters"""
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def save_attachments(service, msg_id, parts, folder):
    if not os.path.exists(folder):
        os.makedirs(folder)

    saved_files = []
    for part in parts:
        if part.get('filename') and 'attachmentId' in part['body']:
            attachment_id = part['body']['attachmentId']
            attachment = service.users().messages().attachments().get(
                userId='me', messageId=msg_id, id=attachment_id).execute()
            file_data = base64.urlsafe_b64decode(attachment['data'])
            file_path = os.path.join(folder, part['filename'])
            with open(file_path, 'wb') as f:
                f.write(file_data)
            saved_files.append(part['filename'])
    return saved_files

def fetch_unread_emails(service, max_results=NUM_EMAILS, sender_name=None):
    query = "in:inbox"
    if sender_name:
        query += f' from:{sender_name}'

    results = service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
    messages = results.get('messages', [])
    email_data = []

    for msg in messages:
        message = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()
        headers = message['payload']['headers']
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '(No Subject)')
        sender = next((h['value'] for h in headers if h['name'] == 'From'), '(Unknown)')
        date = next((h['value'] for h in headers if h['name'] == 'Date'), '(Unknown Date)')
        body = get_message_body(message['payload'])

        email_data.append({
            'From': sender,
            'Subject': subject,
            'Date': date,
            'Body': body,
            'Attachments': [],  # Not used directly
            'Message-ID': msg['id'],
            'Parts': message['payload'].get('parts', [])
        })

    return email_data
# ------------------------------------------
# CLEAN HTML CONTENT (for email body)
# ------------------------------------------
def clean_html_content(html_content):
    """Converts HTML content to plain text or simple formatting"""
    # Create an HTML parser instance
    clean_text = html.unescape(html_content)
    
    # Replace specific HTML tags with simple text formatting (you can expand this as needed)
    clean_text = re.sub(r'<b>(.*?)</b>', r'<font name="Helvetica-Bold">\1</font>', clean_text)
    clean_text = re.sub(r'<i>(.*?)</i>', r'<font name="Helvetica-Oblique">\1</font>', clean_text)
    clean_text = re.sub(r'<u>(.*?)</u>', r'<u>\1</u>', clean_text)
    clean_text = re.sub(r'<br\s*/?>', r'\n', clean_text)  # Convert line breaks to newlines

    # Remove other HTML tags that we don't need
    clean_text = re.sub(r'<[^>]+>', '', clean_text)
    
    return clean_text


# ------------------------------------------
# Email formatter
# ------------------------------------------

def format_email_body_with_llm(email_body):
    api_key = "your-api-key" #openrouter api key

    headers = {
        "Authorization": f"Bearer {api_key}",
        "HTTP-Referer": "http://localhost",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [
            {"role": "system", "content": "You are a helpful assistant that formats raw email text using Markdown to clearly show paragraphs, bold/italic text, headers, and bullet points. Dont add any markdown #"},
            {"role": "user", "content": email_body}
        ]
    }

    response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)

    if response.status_code == 200:
        result = response.json()
        return result['choices'][0]['message']['content']
    else:
        print("Error:", response.status_code, response.text)
        return email_body





    
def convert_markdown_to_pdf(markdown_text, output_filename="email_summary.pdf"):
    html = markdown.markdown(markdown_text)

    with open(output_filename, "w+b") as pdf_file:
        pisa_status = pisa.CreatePDF(html, dest=pdf_file)

    if pisa_status.err:
        print("❌ Error during PDF generation")
    else:
        print(f"✅ PDF generated successfully: {output_filename}")


def summarize_email_body_with_llm(email_body):
    api_key = "your-api-key" #openrouter api key

    headers = {
        "Authorization": f"Bearer {api_key}",
        "HTTP-Referer": "http://localhost",
        "Content-Type": "application/json"
    }

    # Revised prompt to focus on specific sections for the summary
    payload = {
        "model": "mistralai/mistral-7b-instruct",
        "messages": [
            {"role": "system", "content": "You are an assistant trained to summarize emails clearly. Provide the summary in three distinct parts: 1) What is talked about, 2) Key points discussed, and 3) Next steps or actions required. Make sure each part is concise and informative."},
            {"role": "user", "content": email_body}
        ]
    }

    response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)

    if response.status_code == 200:
        result = response.json()
        return result['choices'][0]['message']['content']
    else:
        print("Error:", response.status_code, response.text)
        return "Error generating summary"


def generate_summary_pdf_with_llm(email_body, folder_path):
    """Generate a summary PDF based on the email body formatted by LLM."""
    # Get the summarized content of the email body
    summarized_body = summarize_email_body_with_llm(email_body)

    # Create the summary markdown content
    md_text = f"""
**Summary:**

{summarized_body}
    """

    # Convert the markdown to HTML
    html = markdown.markdown(md_text)

    # Path for the summary PDF
    summary_pdf_filename = os.path.join(folder_path, "email_summary.pdf")

    # Create the PDF
    with open(summary_pdf_filename, "w+b") as pdf_file:
        pisa_status = pisa.CreatePDF(html, dest=pdf_file)

    if pisa_status.err:
        print("❌ Error during Summary PDF generation")
    else:
        print(f"✅ Summary PDF generated successfully: {summary_pdf_filename}")

    return summary_pdf_filename


def generate_email_pdf_with_llm(emails, service, base_folder="email_reports", csv_filename="email_links.csv"):
    if not os.path.exists(base_folder):
        os.makedirs(base_folder)

    with open(csv_filename, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Email", "Subject", "Raw PDF Link", "Summary PDF Link", "Attachment Links"])

        for idx, email in enumerate(emails):
            email_body = email['Body']

            # Generate raw email PDF
            formatted_md_body = format_email_body_with_llm(email_body)
            sender_clean = sanitize_folder_name(email['From'].split('<')[0].strip())
            folder_name = f"email_{idx+1}_{sender_clean}"
            folder_path = os.path.join(base_folder, folder_name)
            os.makedirs(folder_path, exist_ok=True)

            md_text = f"""
**From:** {email['From']}  
**Subject:** {email['Subject']}  
**Date:** {email['Date']}

---

{formatted_md_body}
            """
            raw_pdf_filename = os.path.join(folder_path, "email_raw.pdf")
            convert_markdown_to_pdf(md_text, output_filename=raw_pdf_filename)

            # Upload raw PDF to Dropbox
            dropbox_raw_pdf_path = f"{DROPBOX_BASE_PATH}/{folder_name}/email_raw.pdf"
            raw_pdf_link = upload_to_dropbox(raw_pdf_filename, dropbox_raw_pdf_path)

            # Generate summary PDF
            summary_pdf_filename = generate_summary_pdf_with_llm(email_body, folder_path)

            # Upload summary PDF to Dropbox
            dropbox_summary_pdf_path = f"{DROPBOX_BASE_PATH}/{folder_name}/email_summary.pdf"
            summary_pdf_link = upload_to_dropbox(summary_pdf_filename, dropbox_summary_pdf_path)

            # Save attachments and upload to Dropbox
            msg_id = email['Message-ID']
            parts = email['Parts']
            attachment_files = save_attachments(service, msg_id, parts, folder=folder_path)

            attachment_links = []
            for file in attachment_files:
                file_path = os.path.join(folder_path, file)
                dropbox_file_path = f"{DROPBOX_BASE_PATH}/{folder_name}/{file}"
                file_link = upload_to_dropbox(file_path, dropbox_file_path)
                attachment_links.append(file_link)

            # Write details to CSV
            writer.writerow([email['From'], email['Subject'], raw_pdf_link, summary_pdf_link, ", ".join(attachment_links)])


# ------------------------------------------
# MAIN FUNCTION
# ------------------------------------------
def main():
    service = authenticate_gmail()
    print("✅ Authenticated with Gmail")

    sender_name = input("🔍 Enter sender name to search (leave blank for all unread): ").strip() or None
    emails = fetch_unread_emails(service, sender_name=sender_name)

    if not emails:
        print("📭 No matching unread emails found.")
    else:
        print("✅",len(emails), "email(s) found extracting them")

        print("📨 Generating PDF report...")
        generate_email_pdf_with_llm(emails,service)

if __name__ == "__main__":
    main()
