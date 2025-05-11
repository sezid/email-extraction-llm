# 📧 Email Summarizer with Dropbox Integration

This tool extracts and summarizes emails from **Gmail** and **Yahoo Mail**, converts them into PDFs using Llama's **Mistral  AI (via OpenRouter)**, uploads them to **Dropbox**, and logs metadata (subject, sender, date, Dropbox link) in a CSV file.

---

# ✉️ Gmail Setup

## 1. Enable Gmail API via Google Cloud Console

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create or select a project.
3. Navigate to **APIs & Services > Library**.
4. Search for **Gmail API**, click on it, and click **Enable**.
5. Navigate to **OAuth consent screen**:
   - Select **External**.
   - Fill out app name and user support email.
   - Save and continue through all steps.
6. Navigate to **Credentials**:
   - Click **Create Credentials > OAuth Client ID**.
   - Choose **Desktop App**.
   - Download the `credentials.json` file.

### 2. Install Required Libraries

```bash
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib requests reportlab markdown xhtml2pdf dropbox
```
# ✉️ Yahoo Mail Setup Guide

This guide explains how to configure Yahoo Mail access via IMAP using a secure **App Password** to fetch and summarize emails in your application.

---

## ✅ Prerequisites

- A valid Yahoo Mail account.
- Python installed with basic IMAP libraries (`imaplib`, `email`, etc.).
- Internet connection.

---

## 🔐 Step 1: Enable IMAP in Yahoo Mail

1. Log in to [Yahoo Mail](https://mail.yahoo.com).
2. Click ⚙️ **Settings** (top right) > **More Settings**.
3. Go to the **Mailboxes** tab.
4. Select your email account.
5. Ensure **IMAP** is enabled. (This is usually enabled by default.)

---

## 🔐 Step 2: Generate Yahoo App Password

Yahoo does **not** allow IMAP access with your main email password if 2FA is enabled. Use an **App Password** instead.

1. Visit [Yahoo Account Security](https://login.yahoo.com/account/security).
2. Scroll down to **"App Password"** and click **Generate app password**.
3. Select **Other App** from the dropdown.
4. Type a custom name like `Email Summarizer`, and click **Generate**.
5. Yahoo will display a 16-character app password.
6. Copy it and **save it securely**.

---

## 🛠️ Step 3: Install these libraries

```bash
pip install imapclient email markdown xhtml2pdf requests dropbox
```


## 🔐 Dropbox Access Token Setup

To upload email summaries and attachments to Dropbox, you'll need to create an **Access Token** with specific permissions.

### 📌 Step-by-Step Instructions

1. Go to the [Dropbox App Console](https://www.dropbox.com/developers/apps).
2. Click **Create App**.
3. Choose the following settings:
   - **Scoped Access**
   - **Full Dropbox** access
4. Give your app a unique name and click **Create App**.

---

### ✅ Configure Required Permissions

Go to the **Permissions** tab and enable the following scopes:

- `files.content.read` ✅  
  _Read content from Dropbox (e.g., checking if a file exists)._

- `files.content.write` ✅  
  _Write files to Dropbox (upload summaries and attachments)._

- `sharing.read` ✅  
  _Read shared links to verify access._

- `sharing.write` ✅  
  _Create shared links for PDFs and attachments._

Click **Submit** after enabling the above scopes.

---

### 🔑 Generate Your Access Token

1. Go to the **Settings** tab of your app.
2. Scroll to **OAuth 2** section.
3. Click **Generate Access Token**.
4. Copy the token and save it securely.

---

### Now you are ready to run both these code
