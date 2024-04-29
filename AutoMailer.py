import win32com.client as win32
import os
import json
import comtypes.client

MYCOMP_DOMAIN = "@mycomp.com"
MYCOMP = f"abc{MYCOMP_DOMAIN};cba{MYCOMP_DOMAIN};" \
         f"adc{MYCOMP_DOMAIN};efg{MYCOMP_DOMAIN}; "
HTML_TEXT = """<p>Dear {name1},<br>
                <br>
       Please see {quarter} Quarter {year} calculations of {company} Corp.<br>
       Kindly give us your confirmation<br>
       <br>
       Thanks in advance<br>
       <br>
       <br>
       Kind regards,</p>
       """


def convert_doc_to_pdf(doc_path, conv_list: list, pdf_path=None):
    """
    Convert a DOC file to PDF format.

    Args: doc_path (str): Path to the input DOC file. pdf_path (str): Path to save the output PDF file. If not
    provided, the PDF file will be saved in the same directory with the same name as the
    input file but with the .docx extension.
    :param pdf_path: PDF path parameter if thought to be saved elsewhere
    :param doc_path: Doc file path
    :param conv_list: Already converted log
    """
    if not os.path.exists(doc_path):
        print(f"Error: File '{doc_path}' does not exist.")
        return

    if pdf_path is None:
        pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
    if os.path.exists(pdf_path):
        print("There is already converted file with same name in dir.")
        conv_list.append(pdf_path)
        return
    # Start a new instance of Word
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # Hide the Word application

    # Open the DOC file
    doc = word.Documents.Open(doc_path)

    # Save as DOCX
    doc.SaveAs(pdf_path, FileFormat=17)  # 16 stands for DOCX format

    # Close the documents and the Word application
    doc.Close()
    word.Quit()

    print(f"DOC converted to DOCX: '{pdf_path}'")


class AccountsMailer:
    def __init__(self, path: str, company: str, quarter: int, year: int, company_json: str, attach=True):
        # Initialize the AccountsMailer object with specified parameters
        self.path = os.path.abspath(path)
        self.alrd_conv = []
        self.company = company  # Mail recipient company
        self.quarter = quarter  # Mail quarter
        self.year = year  # Mail year
        self.html_text = HTML_TEXT
        self.attach = attach  # Boolean value to determine if there will be attachments or not
        self.outlook = win32.Dispatch('outlook.application')
        self.mycomp_cc = MYCOMP  # My company cc list
        self.company_json = company_json
        self.reinsurers = self.load_companies()  # Load reinsurers' details from the specified JSON file

    def load_companies(self):
        # Load reinsurers' details from a JSON file and return as a dictionary
        with open(self.company_json, 'r') as f:
            return json.load(f)

    def list_pdf_files(self):
        # Retrieve a list of all PDF files in the specified directory and its subdirectories
        pdf_list = []
        for root, dirs, files in os.walk(self.path):
            for name in files:
                if name.endswith(".pdf"):
                    pdf_list.append(os.path.join(root, name))
                    print(name, "from ", root, "is added to the pdf list.")
        return pdf_list

    def convert(self):
        # Convert all .doc files in the specified directory and its subdirectories to PDF
        for root, dirs, files in os.walk(self.path):
            for file in files:
                if file.endswith(".doc"):
                    file_abs = os.path.join(root, file)
                    convert_doc_to_pdf(file_abs, self.alrd_conv)  # Defined Seperately from class

    def create_email(self, recipient, cc_address, corpname, name=""):
        # Create and configure an email with specified details and send it using Outlook
        name1 = "Madam/Sir" if name == "" else name
        quarters = {
            1: "1st",
            2: "2nd",
            3: "3rd",
            4: "4th"
        }
        # Prepare HTML content for the email
        html_text = self.html_text.format(name1=name1, quarter=quarters[self.quarter], year=self.year,
                                          company=self.company)
        # Create a new email item
        mail = self.outlook.CreateItem(0)
        mail.To = recipient
        mail.CC = self.mycomp_cc + cc_address
        mail.Subject = f"{self.company.capitalize()} CORP {self.quarter}Q{self.year} Accounts - {corpname}"
        # Insert the HTML content into the email body
        index = mail.HTMLBody.find('>', mail.HTMLBody.find('<body'))
        mail.HTMLBody = mail.HTMLBody[:index + 1] + html_text + mail.HTMLBody[index + 1:]
        # Attach PDF files to the email if attachment is enabled
        if self.attach:
            pdflist = self.list_pdf_files()
            for attachment in pdflist:
                if corpname in attachment:
                    mail.Attachments.Add(attachment)
                    print(f"{attachment} is added to {corpname} mail")
        # Display the email (not sending it immediately)
        mail.Display(False)

    def create_mails_all(self):
        # Create emails to each corporate listed in the loaded corp details.
        for corp, details in self.reinsurers.items():
            if corp in os.listdir(self.path):  # Check if directory for the reinsurer exists.
                self.create_email(details["recipient"], details["CC"], corp, details.get("Name", ""))
            else:
                print(f"Directory for reinsurer '{corp}' not found.")
