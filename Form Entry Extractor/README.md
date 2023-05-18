# Form Entry Extractor

This project is capable of extracting form entry details from
notification emails coming from youngskilledindia.com to an email account

## Usage

Go to config file and set the input variables as needed
Run the main.py file to extract emails and update them in excel sheet

## How this software works?

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
All 4 details are stored in a list, which gets stored in a tuple with the creation
time for that email which gets stored in a list of details of all entries.

The list of all entry details are used to update the excel file.
The last updated email in the output file is checked, and matched to get the
index of where to start from in the list of entries fetched online

## Project Terminology

1. message response - complete message response from Gmail API for a particular message
2. message resource - a response containing info about message id etc. but not the content,
   obtained from messages.list in Gmail API
3. message content - message body

## Experience from working on this project

(*) Accessing Gmail using imaplib works with google account password only when less secure
apps are enabled
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(email_address, password)

## Which type of emails does it work with?

It works with emails coming from youngskilledindia.com. 'Email Example.png' is an example
of such email