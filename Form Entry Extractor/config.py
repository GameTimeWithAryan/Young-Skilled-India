"""Config File for entire project"""
from logging import getLogger, StreamHandler, Formatter, INFO, WARNING, DEBUG

# Input Variables
LAST_N_DAYS: int = 1  # Get all form entries from last n days
LOG_LEVEL: int = DEBUG  # INFO, WARNING, DEBUG from logging to show software progress
GMAIL_LABEL_NAME: str = "Form Entries"  # Label from which to extract emails

###############################################################################

# Output Variables
OUTPUT_FILE = "Form Entries.xlsx"
NAME_COLUMN = "A"
EMAIL_COLUMN = "B"
PHONE_COLUMN = "C"
CITY_COLUMN = "D"
DATE_COLUMN = "E"
ENTRY_FIELD_COLUMNS = (NAME_COLUMN, EMAIL_COLUMN, PHONE_COLUMN, CITY_COLUMN, DATE_COLUMN)

# Software Variable
logger = getLogger(__name__)
formatter = Formatter("{%(levelname)s} %(message)s")
steam_handler = StreamHandler()

steam_handler.setFormatter(formatter)
logger.addHandler(steam_handler)
logger.setLevel(LOG_LEVEL)
