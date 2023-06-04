import json
from logging import getLogger, Formatter, StreamHandler, INFO

# Excel data extraction
certificate_data_file: str = "certificate_data.xlsx"  # Certificates data filename
starting_row_index: int = 2  # row index from which to start
column_range: tuple[int, int] = (1, 4)  # columns indices to extract data from

# Email message content
email_subject = "subject"
email_body = """email_body"""

# Gmail login credentials
sender_email: str = "info@youngskilledindia.com"
with open("credentials.json") as cred:
    app_password: str = json.load(cred)["app_password"]  # Gmail App Password

# Software Variables
logger = getLogger(__name__)
formatter = Formatter("{%(levelname)s} %(message)s")
steam_handler = StreamHandler()

steam_handler.setFormatter(formatter)
logger.addHandler(steam_handler)
logger.setLevel(INFO)
