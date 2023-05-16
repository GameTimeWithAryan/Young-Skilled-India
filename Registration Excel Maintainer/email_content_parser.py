"""
Works with google account password only when less secure apps are enabled
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(email_address, password)

Check Gmail API docs for more info
message content is usaually encoded in base64 and quoted printable text encoding

Algorithm for getting name, email, phone number, city -
Seperate the string at \r\n
Get the index of name from list
The next index will be the value of name, next will be email...... DONE!

Word meaning:
message - complete message response, check Gmail API docs for more info
message resource - a response containing info about message id etc. but not the content, obtained from labels.list
message content - message body
"""

import base64
from datetime import date, timedelta
import quopri
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GmailReader:
    def __init__(self, service):
        # Google service resource for interacting with Gmail API
        self.service = service

    def get_label_id(self, label_name: str):
        label_id = None
        # Fetching list of all labels
        label_result: dict = self.service.users().labels().list(userId="me").execute()
        labels: list[dict] = label_result.get("labels")

        # Finding label id
        for label in labels:
            if label.get("name") == label_name:
                label_id = label.get("id")
                break
        return label_id

    def get_msg_resources_under_label(self, label_ids: list[str] | str, custom_filter: str = ""):
        # Using label id to get all message resources from that label
        msg_resources_result: dict = self.service.users().messages().list(userId="me", labelIds=label_ids,
                                                                          q=custom_filter).execute()
        msg_resources: list[dict] = msg_resources_result.get("messages")
        return msg_resources

    def get_msg_from_resource(self, resource: dict):
        message_id = resource.get("id")
        message_result: dict = self.service.users().messages().get(userId="me", id=message_id).execute()
        return message_result

    @staticmethod
    def decode_msg_content(content: str):
        """
        Decodes the encoded content of message given by Gmail API
        First decodes the base64 encoding
        Secondly decodes the Quoted Printable encoding
        """

        # Step 1 - Decoding the base64 text
        b64_decoded_message_content: bytes = base64 \
            .urlsafe_b64decode(content)
        # Step 2 - Decoding the quoted printable text, and converting bytes to str
        quopri_decoded_message_content: str = quopri \
            .decodestring(b64_decoded_message_content) \
            .decode(errors="replace")
        return quopri_decoded_message_content


def parse_content(content: str):
    """Check module docstring for details on how the alogrithm works

    Returns
    -------
    list[name: str, email: str, phone: str, city: str] : a single entry list
    """

    parsed_details = []
    lines_in_content = content.split("\r\n")
    name_field_index = lines_in_content.index("*Name*")
    for index in range(0, 4):
        parsed_details.append(lines_in_content[name_field_index + 1 + (index * 2)])
    return parsed_details


def get_after_date_filter(n_days):
    """Returns today's date - n_days
    To be used in the filter for getting all the emails after a given date
    Return format:
    \"year/month/day\"
    """
    today = date.today()
    after_date = today - timedelta(days=n_days)
    after_date_filter = f"after:{after_date.year}/{after_date.month}/{after_date.day}"
    return after_date_filter


def get_credentials():
    scopes = ['https://www.googleapis.com/auth/gmail.readonly']

    creds = None
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


def get_entry_details(label_name: str, n_days=1):
    """

    Parameters
    ----------
     label_name : str
        Label from which to extract messages
     n_days : int
        Get entry details for all the emails received within last n days

    Returns
    -------
    list[list] : a list of all entries

    """

    all_entry_detais: list = []
    custom_filter = get_after_date_filter(n_days)

    # Get credentials for authentication
    creds = get_credentials()
    # Building Gmail API Caller
    service = build('gmail', 'v1', credentials=creds)
    gmail_reader = GmailReader(service)

    # Getting label id
    label_id = gmail_reader.get_label_id(label_name)
    # Getting all message resources under label
    message_resources: list[dict] = gmail_reader.get_msg_resources_under_label(label_id, custom_filter)

    if message_resources is None:
        print(f"No Form Entry Emails in last {n_days} days")
        return []

    # Getting message content of each message resource
    for index, msg_resources in enumerate(message_resources):
        # Getting each message
        message = gmail_reader.get_msg_from_resource(msg_resources)
        # Extracting content of message, encoded in base64 and quoted printable text encoding
        encoded_message_content: str = message.get("payload").get("parts")[0].get("body").get("data")
        # Decoding message content
        decoded_message_content = gmail_reader.decode_msg_content(encoded_message_content)
        # Parsing content
        entry_details = parse_content(decoded_message_content)
        all_entry_detais.append(entry_details)
    return all_entry_detais


# ENTRY LIST FIELD INDEXES


if __name__ == '__main__':
    print(get_entry_details("Form entries", 2))
