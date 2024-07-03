import os
import base64
import pickle
from email.mime.text import MIMEText
from flask import Flask, request, render_template, redirect, url_for, flash
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Replace with your secret key
app.config['UPLOAD_FOLDER'] = 'uploads'

SCOPES = ['https://www.googleapis.com/auth/gmail.send']
SHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_emails', methods=['POST'])
def send_emails():
    sheet_id = request.form['sheet_id']
    sender_email = request.form['sender_email']
    subject = request.form['subject']

    credentials_file = request.files['credentials']
    credentials_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(credentials_file.filename))
    credentials_file.save(credentials_path)

    sheets_service = get_sheets_service(credentials_path)
    df = data_extractor(sheets_service, sheet_id)

    if df.empty:
        return "No data to send emails."

    gmail_service = get_gmail_service(credentials_path)
    for index, row in df.iterrows():
        email = row['email']
        body = row['Message']
        send_email(gmail_service, email, subject, body, sender_email)

    os.remove(credentials_path)
    return "Emails sent successfully!"

def get_gmail_service(credentials_path):
    creds = None
    if os.path.exists('token_gmail.pickle'):
        with open('token_gmail.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token_gmail.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

def get_sheets_service(credentials_path):
    creds = None
    if os.path.exists('token_sheets.pickle'):
        with open('token_sheets.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path, SHEET_SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token_sheets.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service

def send_email(service, to, subject, body, sender):
    message = MIMEText(body)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    message = {'raw': raw}
    try:
        sent_message = service.users().messages().send(userId='me', body=message).execute()
        print('Message Id: %s' % sent_message['id'])
    except Exception as error:
        print(f'An error occurred: {error}')

def data_extractor(service, sheet_id):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range='Sheet1').execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
        return pd.DataFrame()
    else:
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df

if __name__ == '__main__':
    app.run(debug=True)
