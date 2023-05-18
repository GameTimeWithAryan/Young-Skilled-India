"""Config File for entire project"""

# Input Variables
GMAIL_LABEL_NAME: str = "Form Entries"  # Label from which to extract emails
LAST_N_DAYS: int = 1  # Get all form entries from last n days

# Output Variables
OUTPUT_FILE = "Form Entries.xlsx"
NAME_COLUMN = "A"
EMAIL_COLUMN = "B"
PHONE_COLUMN = "C"
CITY_COLUMN = "D"
DATE_COLUMN = "E"
ENTRY_FIELD_COLUMNS = (NAME_COLUMN, EMAIL_COLUMN, PHONE_COLUMN, CITY_COLUMN, DATE_COLUMN)
