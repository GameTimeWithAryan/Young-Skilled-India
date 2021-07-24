from email.message import EmailMessage
from openpyxl import load_workbook
from smtplib import SMTP_SSL
import pyperclip
import keyboard
import imghdr
import os


# Loading excel Workbook to get Worksheet for emails and names
WORKBOOK = load_workbook('python data.xlsx')
WORKSHEET = WORKBOOK.active

# Email credentials for gmail SMTP authentication
EMAIL_USERNAME = os.environ.get('YSI_PYTHON_EMAIL_USERNAME')
EMAIL_PASSWORD = os.environ.get('YSI_PYTHON_EMAIL_PASSWORD')

# Reading the value from starting.txt to get from where to start, because if the program crashes for some reason then it
# Can resume from where it left
with open('Starting.txt', 'r') as Names_R_File:
    STARTING = int(Names_R_File.read())
    ENDING = 79

BODY = """\
Hi!
Greetings of the day!

Congratulations! Please find the Participation Certificate for the Management Skills Workshop.

Based on your feedback & interest to  Join the "Management Certificate Course ( Product evaluated by AICTE)" after participating in the Young Skilled India management Skills workshop on 18th Jul'21, to upgrade your career growth path. This is a new initiative to minimize the 2 Years MBA skills in 6 Months with 90% practical learnings, because working professionals may not have time to Join the 2Year Traditional MBA classes regularly, it's a Next-generation industry-recommended course, Time saving and cost-effective.

Please find the method to enroll under the Ministry of Education, NEAT, AICTE Scheme.
Step 1: Please register at NEAT, AICTE website as a LEARNER by clicking the course Link:

Course Link: https://neat.aicte-india.org/course-details/NEAT2020622_PROD_2

Step 2: After successful registration, Kindly click the above link  ( https://neat.aicte-india.org/course-details/NEAT2020622_PROD_2)  again to pay the fees and confirm your admission for the next batch.

Course Duration: 6 Months, Online Live & Interactive Class timing: Saturday & Sunday ( 6 PM to 8 PM)
Faculty: Renowned Professors / Industry Experts
Fee: Rs 18,000/ Only for 6 months course.
Certification after completion: Master Certificate in Business Management with specialization in HR/Marketing/Finance/Operations/SCM/QC/IT etc.
Certification Validity: Globally

Management Certification Course Introductory Session for the previous batch: https://www.youtube.com/watch?v=z3gzEVzS0fU&ab_channel=Youngskilledindia.Com

In case of any assistance please feel free to call at M-8009321506


Regards,
Niraj Srivastava
Founder & CEO
Young Skilled India
(YSIID Solutions Pvt. Ltd, 
A Govt. of India Certified Startup DIPP 1656)
M-8009321506
Incubated at: NCL - IIT BHU MCIIE\
"""


# Utility Functions
def copy_to_clipboard(text, dummy):
    pyperclip.copy(text)


# Get Cells Of An excel file which can be used later to get value of that cell by "cell.value"
# Returns a list of all the cells of a excel file in the specified cell_name, cell_name = column of the cells
def get_excel_cells(cell_name, starting, ending):
    all_cells = []
    for cell_no in range(starting, ending):
        cell = WORKSHEET[cell_name + str(cell_no)]
        all_cells.append(cell)
    return all_cells


# Send the message from my gmail
# param type email.message.EmailMessage
def send_message(msg):
    with SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        smtp.send_message(msg)


# Main Codes For Copy Certificate Names And Mail_Certificate_To_Emails
def Copy_Certificate_Names(Name_List, Starting):
    # Iterating over the Name_List for names
    for name in Name_List:
        cell_name = name.value
        print(f"Next:   {cell_name}")

        # add hotkey requires more than one parameter in "args" set or else it behaves weirdly
        keyboard.add_hotkey('/', copy_to_clipboard, (cell_name, "dummy"))
        keyboard.wait('/')

        print(f'Copied: {cell_name}')

        # Updating the values from the Starting.txt file
        Starting += 1
        with open('Starting.txt', 'w') as W_File:
            W_File.write(str(Starting))


def Mail_Certificate_To_Emails(Email_List, Names_List):
    # Iterate over ever value in the Emails_List for emails
    for index, email in enumerate(Email_List):
        email_id = email.value
        name = Names_List[index].value

        # Email Message Object which is used to send messages
        Message = EmailMessage()
        Message['From'] = EMAIL_USERNAME
        Message['To'] = email_id
        Message['Subject'] = "YSI - Certificate For Management Skills Workshop"

        # Reading binary from the photo to add in the message
        with open("YSI - Certificates/" + name + ".jpg", 'rb') as Image:
            Image_data = Image.read()
            Image_name = os.path.basename(Image.name)
            Image_type = imghdr.what(Image.name)

        # Setting message body and adding the photo attachment and then sending the message
        Message.set_content(BODY)
        Message.add_attachment(Image_data, maintype='image', subtype=Image_type, filename=Image_name)
        send_message(Message)
        print(f"Mail Sent!\nIndex: {index}\n")


if __name__ == "__main__":
    ######################### Copy Certificate Names #########################
    # Names = get_excel_cells("C", STARTING, ENDING)
    # Copy_Certificate_Names(Names, STARTING_NAMES)

    ######################### Sent Certificates To Emails #########################
    # Names = get_excel_cells("C", STARTING, ENDING)
    # Emails = get_excel_cells('B', STARTING, ENDING)
    # Mail_Certificate_To_Emails(Emails, Names)
    pass
