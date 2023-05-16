"""
email_content_parser.py
~~~~~~~~~~~~~~~~~~~~~~~

Entry details extractor for emails coming from youngskilledindia.com forms
The get_entry_detials function gets all the emails from Gmail API, under
the label "Form Entries", and then parses them to extract the text under
those fields, which are the name, email ID, phone number and city of a person


Project Terminology
-------------------
message - complete message response from Gmail API for a particular message
message resource - a response containing info about message id etc. but not the content
                    obtained from messages.list in Gmail API
message content - message body


THINGS TO KNOW WHILE WORKING
----------------------------
(*) Made by reading Gmail API Documentation

(*) Accessing Gmail using imaplib works with google account password when less secure apps are enabled
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(email_address, password)

(*) Message content is encoded in base64 and quoted printable text encoding for
the emails with which I am working

(*) Algorithm for getting name, email, phone number, city
Seperate the string at \r\n
Get the index of name field of email i.e. "*Name*" string from list
The next index will be the value of name field, next will be email field...... city field value, DONE!
"""

import base64
from datetime import date, timedelta, datetime
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


def parse_content(content: str):
    """Parses the contents of the email to obtain name, email, phone and city
    Check module docstring for details on how the alogrithm works

    Returns
    -------
    [name: str, email: str, phone: str, city: str]
        a single entry list
    """

    parsed_details = []
    lines_in_content = content.split("\r\n")
    name_field_index = lines_in_content.index("*Name*")
    for index in range(0, 4):
        parsed_details.append(lines_in_content[name_field_index + 1 + (index * 2)])
    return parsed_details


def get_entry_details(label_name: str, n_days=1):
    """Get entry details for all emails receieved in last `n_days`
    under `label_name` along with time at which it was sent

    Parameters
    ----------
     label_name : str
        Label from which to extract messages
     n_days : int
        Get entry details for all the emails received within last n days

    Returns
    -------
    list[(entry: list, date: str)] : a list of all entries with the time they were sent

    Notes
    -----
    How this module works?
    It uses a google cloud project running Gmail API service to access Gmail.
    It first authenticates using credentials.json file downloaded from
    credentials tab in APIs & Services tab in google cloud console.
    Then it lists the IDs (other metadata are also sent by gmail API but are of
    no interest to us) of all the emails under a particular label and after a
    particular date.
    The IDs are then used to retrieve each email individually with the time
    it was sent along with its content/body which contains the entry detials.
    Each email's content is then extracted from the response, and it's content
    is decoded, as the response by gmail contains base64 and quoted printable
    encoded text.
    The decoded content is then parsed to extract name, email, phone number and
    city name from the emails.
    All 4 details are stored in a list, which gets stored in a list of all
    entries along with the time when they were sent for all entries,
    and returned at the end by this function.
    """

    all_entry_detais: list = []
    # Filter for getting all emails for the label name and after a given date
    custom_filter = f"label:{label_name} {get_after_date_filter(n_days)}"

    # Get credentials for authentication
    creds = get_credentials()
    # Building Gmail API Caller
    service = build('gmail', 'v1', credentials=creds)
    gmail_reader = GmailReader(service)

    # Getting all message resources in the filter
    message_resources: list[dict] = gmail_reader.get_msg_resources_in_filter(custom_filter)
    if message_resources is None:
        print(f"No Form Entry Emails in last {n_days} days")
        return []

    # Getting message content of each message resource
    for index, msg_resource in enumerate(message_resources):
        # Getting each message
        message_response = gmail_reader.get_msg_from_resource(msg_resource)
        # Extracting creation date of message from message_response
        msg_creation_date = get_msg_creation_date(message_response)
        # Extracting content of message, encoded in base64 and quoted printable text encoding
        encoded_message_content: str = message_response.get("payload").get("parts")[0].get("body").get("data")
        # Decoding message content
        decoded_message_content = gmail_reader.decode_msg_content(encoded_message_content)
        # Parsing content
        entry_details = parse_content(decoded_message_content)
        all_entry_detais.append((entry_details, msg_creation_date))

    # Reversing to make the entries in ascending order of their sent dates
    all_entry_detais = list(reversed(all_entry_detais))

    return all_entry_detais


if __name__ == '__main__':
    l = get_entry_details("Form Entries", 1)
    for item, d in l:
        print(f"Entry: {item}, Date: {d}")
