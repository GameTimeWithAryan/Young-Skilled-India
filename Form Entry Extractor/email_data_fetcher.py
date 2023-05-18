"""
email_data_fetcher.py
~~~~~~~~~~~~~~~~~~~~~~~

Content and sent time extractor for emails coming from youngskilledindia.com forms
The yield_emails_content_with_date function gets all the emails from Gmail API,
under a specific label, and then yields its content and its sent time for each
email
"""

from config import GMAIL_LABEL_NAME, LAST_N_DAYS

import base64
import quopri
import os.path
from datetime import date, timedelta, datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GmailReader:
    def __init__(self):
        self.service = None

    def initialize_api_caller(self):
        # Get credentials for authentication
        creds = get_credentials()
        # Building Gmail API Caller
        self.service = build('gmail', 'v1', credentials=creds)

    def get_msg_resources_in_filter(self, custom_filter: str):
        """Gets all message resources for the `custom_filter`"""
        msg_resources_result: dict = self.service.users().messages().list(userId="me", q=custom_filter).execute()
        msg_resources: list[dict] = msg_resources_result.get("messages")
        return msg_resources

    def get_msg_from_resource(self, resource: dict):
        """Gets full message response using the message resource"""
        message_id = resource.get("id")
        message_result: dict = self.service.users().messages().get(userId="me", id=message_id).execute()
        return message_result

    @staticmethod
    def decode_msg_content(content: str):
        """Decodes the encoded content of message given by Gmail API
        First decodes the base64 encoding
        Second decodes the Quoted Printable encoding
        """

        # Step 1 - Decoding the base64 text
        b64_decoded_message_content: bytes = base64 \
            .urlsafe_b64decode(content)
        # Step 2 - Decoding the quoted printable text, and converting bytes to str
        quopri_decoded_message_content: str = quopri \
            .decodestring(b64_decoded_message_content) \
            .decode(errors="replace")
        return quopri_decoded_message_content


def get_credentials():
    creds = None
    scopes = ['https://www.googleapis.com/auth/gmail.readonly']
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def get_after_date_filter(n_days):
    """Returns filter for selecting all emails after the date before `n_days` from today
    To be used in the filter for getting all the emails after a given date"""

    today = date.today()
    after_date = today - timedelta(days=n_days)
    after_date_filter = f"after:{after_date.year}/{after_date.month}/{after_date.day}"
    return after_date_filter


def get_msg_creation_date(message_response):
    msg_creation_milliseconds = message_response.get("internalDate")
    msg_datetime = datetime.fromtimestamp(int(msg_creation_milliseconds) // 1000)
    msg_creation_date = f"{msg_datetime.day}/{msg_datetime.month}/{msg_datetime.year}"
    return msg_creation_date


def yield_email_data(label_name: str = GMAIL_LABEL_NAME, n_days: int = LAST_N_DAYS):
    """Yields the email content and sent date for each email"""
    # Filter for getting all emails for the label name and after a given date
    custom_filter = f"label:{label_name} {get_after_date_filter(n_days)}"
    gmail_reader = GmailReader()
    gmail_reader.initialize_api_caller()

    # Getting all message resources in the filter
    message_resources: list[dict] = gmail_reader.get_msg_resources_in_filter(custom_filter)
    if message_resources is None:
        print(f"No Form Entry Emails in last {n_days} days")
        return []

    # Yielding message content and sent date of each message resource
    for msg_resource in message_resources:
        # Getting each message
        message_response = gmail_reader.get_msg_from_resource(msg_resource)
        # Extracting content of message, encoded in base64 and quoted printable text encoding
        encoded_message_content: str = message_response.get("payload").get("parts")[0].get("body").get("data")
        # Decoding message content
        decoded_message_content = gmail_reader.decode_msg_content(encoded_message_content)
        # Extracting creation date of message from message_response
        msg_creation_date = get_msg_creation_date(message_response)

        yield decoded_message_content, msg_creation_date


def main():
    print("Email Fetcher")
    print("Get email entries of last how many days?")
    n_days = int(input("Days - "))
    for content, creation_date in yield_email_data(n_days=n_days):
        print(content)
        print(creation_date)
        print("-" * 80)


if __name__ == '__main__':
    main()
