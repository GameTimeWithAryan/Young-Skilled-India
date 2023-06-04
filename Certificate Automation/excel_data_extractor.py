from config import certificate_data_file, starting_row_index, column_range

from openpyxl import load_workbook


def yield_excel_row_data():
    # Loading excel Workbook
    workbook = load_workbook(certificate_data_file)
    worksheet = workbook.worksheets[0]
    # Iterating over the rows in worksheet
    for row_index, row in enumerate(worksheet):
        if row_index < starting_row_index - 1:
            continue

        # Yielding data of each row
        row_data = []
        for column in row[column_range[0] - 1: column_range[1]]:
            row_data.append(column.value)
        yield row_data
