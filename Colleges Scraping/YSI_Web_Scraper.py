# Importing required modules
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pyautogui

# BASE Variables
PARSER = 'html.parser'
# URL template to make URLs From The Base URL
TEMPLATE_URL = "http://education-india.in/Education/Colleges/College_Details.php?CollegeId={{COLLEGE_ID}}"
# Making WORKBOOK And WORKSHEET For Later Use
WORKBOOK = load_workbook('College_Data.xlsx')
WORKSHEET = WORKBOOK.active

"""
COLLEGE_ID Variables:
COLLEGE_ID_ENDING - Max Values For 'CollegeId' Param In TEMPLATE_URL
COLLEGE_ID_STARTING- Starting Values For 'CollegeId' Param In TEMPLATE_URL

COLLEGE_ID_STARTING Is Taken From COLLEGE_ID_STARTING.txt File So that if program crashes for some reason then the 
starting value Could be retained And it do not require to start from 1 again

COLLEGE_ID_STARTING Is Read from COLLEGE_ID_STARTING.txt file at the start and its value is updated after every 
iteration in the file
"""
COLLEGE_ID_ENDING = 7141
with open('COLLEGE_ID_STARTING.txt', 'r') as File:
    COLLEGE_ID_STARTING = int(File.read())

# Identifiers For Matching Table Heads To Get To Table Data
NAME_TH = '<th>Name</th>'
ADDRESS_TH = '<th scope="row">Address</th>'
CONTACT_TH = '<th scope="row">Contact</th>'
EMAIL_ID_TH = '<th scope="row">Email ID</th>'
WEBSITE_TH = '<th scope="row">Website</th>'

# CELL Names For Excel Sheet
NAME_CELL = 'A'
ADDRESS_CELL = 'B'
CONTACT_CELL = 'C'
EMAIL_ID_CELL = 'D'
WEBSITE_CELL = 'E'

College_Name = ''
College_Address = ''
College_Contact = ''
College_EmailId = ''
College_Website = ''


# Utility Functions
# Get "next sibling's text" of a beautiful soup object
def get_next_next_sibling_text(Tag):
    return Tag.next_sibling.next_sibling.get_text()


# Add ", " between Pincode, District and State for beauty
def beautify_address(Address):
    Pincode_Index = Address.find('Pincode')
    District_Index = Address.find('District')
    State_Index = Address.find('State')

    Address = Address[:Pincode_Index] + ", " + Address[Pincode_Index:District_Index] + ", " + \
              Address[District_Index:State_Index] + ", " + Address[State_Index:]
    return Address


# Print info of the current college in iteration
def print_info(Name, Address, Contact, Email_ID, Website, College_Number):
    print(f"{Name = }")
    print(f"{Address = }")
    print(f"{Contact = }")
    print(f"{Email_ID = }")
    print(f"{Website = }")
    print(f"{College_Number = }")
    print()


# Add the scraped data to a excel worksheet
def add_to_worksheet(Worksheet, Cell_Name, Cell_Number, Cell_Value):
    Cell = Cell_Name + str(Cell_Number)
    Worksheet[Cell].value = Cell_Value


if __name__ == '__main__':
    for College_Id in range(COLLEGE_ID_STARTING, COLLEGE_ID_ENDING):
        Table_Number = College_Id + 1
        url = TEMPLATE_URL.replace("{{COLLEGE_ID}}", str(College_Id))
        html_content = requests.get(url).content
        soup = BeautifulSoup(html_content, PARSER)
        details_table = soup.find("table", class_='detail')
        table_heads = details_table.find_all('th')

        for table_head in table_heads:
            string_table_head = str(table_head)
            if string_table_head == NAME_TH:
                College_Name = get_next_next_sibling_text(table_head)
                add_to_worksheet(WORKSHEET, NAME_CELL, Table_Number, College_Name)

            elif string_table_head == ADDRESS_TH:
                College_Address = beautify_address(get_next_next_sibling_text(table_head))
                add_to_worksheet(WORKSHEET, ADDRESS_CELL, Table_Number, College_Address)

            elif string_table_head == CONTACT_TH:
                College_Contact = get_next_next_sibling_text(table_head)
                add_to_worksheet(WORKSHEET, CONTACT_CELL, Table_Number, College_Contact)

            elif string_table_head == EMAIL_ID_TH:
                College_EmailId = get_next_next_sibling_text(table_head)
                add_to_worksheet(WORKSHEET, EMAIL_ID_CELL, Table_Number, College_EmailId)

            elif string_table_head == WEBSITE_TH:
                College_Website = get_next_next_sibling_text(table_head)
                add_to_worksheet(WORKSHEET, WEBSITE_CELL, Table_Number, College_Website)

        print_info(College_Name, College_Address, College_Contact, College_EmailId, College_Website, College_Id)
        WORKBOOK.save('College_Data.xlsx')
        Table_Number += 1
        College_Id += 1

        with open('COLLEGE_ID_STARTING.txt', 'w') as File:
            File.write(str(College_Id))

        pyautogui.press('space')
