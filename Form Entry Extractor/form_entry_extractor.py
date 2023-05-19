"""
form_entry_extractor.py
~~~~~~~~~~~~~~~~~~~~~~~

Extracts form entry details from email content

Notes
-----
Algorithm for getting name, email, phone number, city:
(1) Seperate the content string at \r\n
(2) Get the index of name field of email i.e. "*Name*" string from list
(3) The next index will be the value of name field, next will be email field...... city field value, DONE!
"""

from email_data_fetcher import yield_email_data
from config import logger


def parse_content(content: str):
    """Parses the contents of the email to obtain name, email, phone and city
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


def get_all_form_entries():
    """Gets all form entries for all emails receieved in `LAST_N_DAYS`
    under `GMAIL_LABEL_NAME` along with time at which it was sent

    Returns
    -------
    list[(entry: list, date: str)]
        a list of tuples of all entries with the time they were sent
    """

    all_entry_detais: list = []
    all_email_data = yield_email_data()

    # Getting message content of each message resource
    for index, (content, msg_creation_date) in enumerate(all_email_data):
        # Parsing content of each message
        entry_details = parse_content(content)
        logger.info(f"[PARSED] Message {index + 1} Parsed")
        all_entry_detais.append((entry_details, msg_creation_date))

    # Reversing to make the entries in ascending order of their sent dates
    all_entry_detais = list(reversed(all_entry_detais))

    return all_entry_detais


def main():
    for entry_detail in get_all_form_entries():
        print(entry_detail)


if __name__ == '__main__':
    main()
