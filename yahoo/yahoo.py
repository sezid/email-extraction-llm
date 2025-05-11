import imaplib
import email
from email.header import decode_header
import markdown
from xhtml2pdf import pisa
import requests
import dropbox
import csv

DROPBOX_ACCESS_TOKEN = "your-dropbox-token" #your dropbox access token to store emails
DROPBOX_UPLOAD_FOLDER = "/YahooEmails"
CSV_OUTPUT_FILE = "email_summaries.csv"

dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)

# CONFIG
EMAIL_ACCOUNT = "your-email" #your yahoo email
PASSWORD = "your-app-password" #your app password
IMAP_SERVER = "imap.mail.yahoo.com"
IMAP_PORT = 993
NUM_EMAILS = 100

# Email formatter (LLM-based using OpenRouter)
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
            {"role": "system", "content": "You are a helpful assistant that formats raw email text using Markdown to clearly show paragraphs, bold/italic text, headers, and bullet points. Don't add any markdown #"},
            {"role": "user", "content": email_body}
        ]
    }

    response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)

    if response.status_code == 200:
        result = response.json()
        return result['choices'][0]['message']['content']
    else:
        print("❌ LLM formatting failed:", response.status_code, response.text)
        return email_body

# PDF Generator
def convert_markdown_to_pdf(markdown_text, output_filename="email_summary.pdf"):
    html_content = markdown.markdown(markdown_text)

    with open(output_filename, "w+b") as pdf_file:
        pisa_status = pisa.CreatePDF(html_content, dest=pdf_file)

    if pisa_status.err:
        print("❌ PDF generation error")
    else:
        print(f"✅ PDF created: {output_filename}")

def clean(text):
    return "".join(c if c.isalnum() else "_" for c in text)

def decode_mime_words(s):
    decoded = decode_header(s)
    return ''.join([part.decode(encoding or 'utf-8') if isinstance(part, bytes) else part for part, encoding in decoded])

def save_attachments(msg, email_index):
    attachment_links = []
    for part in msg.walk():
        content_disposition = str(part.get("Content-Disposition"))
        if part.get_content_maintype() == 'multipart':
            continue
        if "attachment" in content_disposition:
            filename = part.get_filename()
            if filename:
                filename = decode_mime_words(filename)
                
                # Upload attachment to Dropbox directly
                dropbox_path = f"{DROPBOX_UPLOAD_FOLDER}/attachment_{email_index}_{filename}"
                attachment_link = upload_to_dropbox(part.get_payload(decode=True), dropbox_path)
                attachment_links.append(attachment_link)
    return attachment_links

def upload_to_dropbox(file_data, dropbox_path):
    dbx.files_upload(file_data, dropbox_path, mode=dropbox.files.WriteMode("overwrite"))
    shared_link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_path)
    return shared_link_metadata.url.replace("?dl=0", "?raw=1")  # Direct link to file

# Connect & Fetch Emails
def fetch_emails(search_query="ALL"):
    csv_data = []
    print("🔐 Logging into Yahoo IMAP...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select("inbox")

    result, data = mail.search(None, search_query)
    if result != "OK":
        print("❌ No messages found.")
        return

    email_ids = data[0].split()
    print(f"📩 Found {len(email_ids)} email(s).")

    for i, email_id in enumerate(email_ids, start=1):
        result, msg_data = mail.fetch(email_id, "(RFC822)")
        if result != "OK":
            print(f"⚠️ Could not fetch email {i}")
            continue

        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_mime_words(msg.get("Subject", "No Subject"))
        sender = msg.get("From", "Unknown Sender")
        date = msg.get("Date", "Unknown Date")

        print(f"\n📨 Email {i}")
        print(f"   From: {sender}")
        print(f"   Subject: {subject}")
        print(f"   Date: {date}")

        # Extract plain text body
        body = None
        html_body = None

        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))

            if part.get_content_maintype() == "multipart":
                continue

            payload = part.get_payload(decode=True)
            charset = part.get_content_charset() or "utf-8"

            try:
                decoded = payload.decode(charset, errors="replace")
            except Exception as e:
                decoded = payload.decode("utf-8", errors="replace")

            if content_type == "text/plain" and "attachment" not in content_disposition:
                body = decoded
                break  # Prefer plain text and stop
            elif content_type == "text/html" and html_body is None:
                html_body = decoded  # Save HTML as fallback

        if not body:
            if html_body:
                import bs4
                soup = bs4.BeautifulSoup(html_body, "html.parser")
                body = soup.get_text(separator="\n")
            else:
                body = "[No plain text or HTML body found]"


        # Use LLM to format and convert to PDF
        print("🧠 Formatting email with LLM...")
        formatted_body = format_email_body_with_llm(body)

        markdown_text = f"""
        **From:** {sender}  
        **Subject:** {subject}  
        **Date:** {date}  

        --- 

        {formatted_body}
        """

        pdf_path = f"email_summary_{i}.pdf"
        convert_markdown_to_pdf(markdown_text, pdf_path)

        # Upload PDF to Dropbox and get the link
        print(f"📄 Uploading PDF to Dropbox...")
        dropbox_pdf_path = f"{DROPBOX_UPLOAD_FOLDER}/email_{i}_{sender}.pdf"
        dropbox_pdf_link = upload_to_dropbox(open(pdf_path, 'rb').read(), dropbox_pdf_path)
        print(f"☁️ Dropbox PDF link: {dropbox_pdf_link}")

        # Save attachments and get their links
        attachment_links = save_attachments(msg, i)

        # Append email data to CSV list
        csv_data.append({
            "Subject": subject,
            "From": sender,
            "Date": date,
            "Dropbox PDF Link": dropbox_pdf_link,
            "Attachment Links": "; ".join(attachment_links)  # List of attachment links
        })

    mail.logout()

    # Save CSV data after processing all emails
    if csv_data:
        with open(CSV_OUTPUT_FILE, mode="w", newline='', encoding="utf-8") as csv_file:
            fieldnames = ["Subject", "From", "Date", "Dropbox PDF Link", "Attachment Links"]
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()
            for row in csv_data:
                writer.writerow(row)
        print(f"📄 CSV summary saved: {CSV_OUTPUT_FILE}")

# MAIN
def main():
    print("📬 Yahoo Mail Reader with Search & Attachments")
    keyword = input("🔍 Enter keyword to search (blank for all): ").strip()

    if keyword:
        search_query = f'TEXT "{keyword}"'
    else:
        search_query = "ALL"

    fetch_emails(search_query)

if __name__ == "__main__":
    main()
