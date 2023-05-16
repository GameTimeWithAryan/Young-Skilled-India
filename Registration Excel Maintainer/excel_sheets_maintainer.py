from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from email_content_parser import get_entry_details

# IMPORTANT Variables
filename = "Form Entries.xlsx"
last_n_days_entries = 3  # Get all form entries of last n days
gmail_label_name = "Form entries"  # Label name in Gmail

# CONSTANTS
NAME_COLUMN = "A"
EMAIL_COLUMN = "B"
PHONE_COLUMN = "C"
CITY_COLUMN = "D"
ENTRY_FIELD_COLUMNS = (NAME_COLUMN, EMAIL_COLUMN, PHONE_COLUMN, CITY_COLUMN)


def main():
    # Loading excel workbook
    workbook = load_workbook(filename)
    worksheet: Worksheet = workbook.worksheets[0]

    form_entries: list = get_entry_details(gmail_label_name, last_n_days_entries)
    last_row: int = worksheet.max_row  # Last updated row
    last_updated_email: str = worksheet[f"{EMAIL_COLUMN}{last_row}"].value

    # Getting index from which entry to start adding in the worksheet
    # If the last updated entry is in the list, then start adding entries
    # from the entry after the last updated entry, identified by its email
    # If the last updated entry is not in the list, then add all the entries
    list_starting_index = 0
    for index, entry in enumerate(form_entries):
        if entry[1] == last_updated_email:
            list_starting_index = index + 1

    # Adding each entry to excel worksheet
    for entry_index, entry in enumerate(form_entries[list_starting_index:]):
        for field_index, field in enumerate(entry):
            cell_coordinate = f"{ENTRY_FIELD_COLUMNS[field_index]}{last_row + entry_index + 1}"
            worksheet[cell_coordinate] = entry[field_index]

    workbook.save(filename=filename)


if __name__ == "__main__":
    main()
