"""
Works with google account password only when less secure apps are enabled
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(email_address, password)

Gmail API docs
"""
import base64
import quopri
import os.path
from pprint import pprint

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def get_credentials():
    # If modifying these scopes, delete the file token.json.
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


def main():
    test_label_id = None
    test_label_name = "Test, using quora"

    try:
        # Get credentials for authentication
        creds = get_credentials()
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        # Fetching label result from Gmail API
        label_result: dict = service.users().labels().list(userId="me").execute()
        labels: list[dict] = label_result.get("labels")

        # Getting label id of test label
        for label in labels:
            if label.get("name") == test_label_name:
                test_label_id = label.get("id", None)
                break

        # Using label id to get all messages from test label
        messages_result: dict = service.users().messages().list(userId="me", labelIds=test_label_id).execute()
        messages: list[dict] = messages_result.get("messages")
        # Getting additional message details by using get method on each message
        for index, message in enumerate(messages):
            # Getting each message from API
            message_id = message.get("id")
            complete_message: dict = service.users().messages().get(userId="me", id=message_id).execute()

            # Getting content of message, quoted printable text, in base64 encoding
            encoded_message_content: str = complete_message.get("payload").get("parts")[0].get("body").get("data")

            # Decoding message content
            # Step 1 - Decoding base64
            b64_decoded_message_content: bytes = base64 \
                .urlsafe_b64decode(encoded_message_content)
            # Step 2 - Decoding quoted printable text, and converting bytes to str
            quopri_decoded_message_content: str = quopri \
                .decodestring(b64_decoded_message_content) \
                .decode(errors="replace")

            # Parsing content
            pass

    except HttpError as error:
        print(f'An HTTP error occurred: {error}')


if __name__ == '__main__':
    main()
