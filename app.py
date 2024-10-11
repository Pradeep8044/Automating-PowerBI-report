import os
import base64
import pickle
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from datetime import datetime, timedelta

# Define the scope for Gmail API
SCOPES = []

def authenticate_gmail():
    creds = None
    # Token is created during the first run, reuse it for authentication
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials are available or token is expired, refresh or get a new one
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())  # Refresh the token
            except Exception as e:
                print(f"Token refresh failed: {e}. Starting fresh login.")
                creds = None  # Force a new login if refresh fails
        if not creds:  # If no valid credentials, log in again
            flow = InstalledAppFlow.from_client_secrets_file('key', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for future use
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def get_yesterdays_date():
    """Returns yesterday's date in the format used in the email subject."""
    return (datetime.now() - timedelta(1)).strftime("%b %d")

def search_email(service, yesterday):
    """Search for emails with the sales report from yesterday."""
    query = f'subject:Daily Sales Report {yesterday}, 2024'
    results = service.users().messages().list(userId='me', q=query).execute()
    messages = results.get('messages', [])
    return messages[0]['id'] if messages else None

def download_attachment(service, message_id):
    """Download the 'Sales_History' attachment from the identified email."""
    message = service.users().messages().get(userId='me', id=message_id).execute()
    for part in message['payload'].get('parts', []):
        if part['filename'] and 'Sales_History' in part['filename'] and 'csv' in part['filename']:
            attachment_id = part['body']['attachmentId']
            attachment = service.users().messages().attachments().get(userId='me', messageId=message_id, id=attachment_id).execute()
            data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            return part['filename'], data
    return None, None

def save_csv(data):
    """Replace the existing file used in Power BI with the new sales report."""
    # The file path where Power BI looks for the CSV (constant filename)
    file_path = r'C:/Users/prade/OneDrive/Desktop/MODEL/New folder/Sales_History_Sep_18,_2024.csv'
    
    # Save the new data, overwriting the old file
    with open(file_path, 'wb') as f:
        f.write(data)
    print('Sales report CSV replaced successfully.')

def main():
    try:
        service = authenticate_gmail()

        # Get yesterday's date for the email subject
        yesterday = get_yesterdays_date()

        # Search for the email with yesterday's date
        email_id = search_email(service, yesterday)

        if email_id:
            filename, data = download_attachment(service, email_id)
            if filename and data:
                save_csv(data)  # Save the CSV and replace the old one
            else:
                print('No Sales_History CSV attachment found.')
        else:
            print(f'No email found for {yesterday}.')
    
    except Exception as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    main()
