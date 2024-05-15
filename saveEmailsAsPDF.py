import os
import shutil
import win32com.client
from win32com.client import constants
from getpass import getpass
from fpdf import FPDF
from datetime import datetime


def find_pdf_with_prefix(folder_path, prefix):
    """
    Search for a PDF file in the specified folder with the given 8-character prefix.

    :param folder_path: Path of the folder to search in.
    :param prefix: The 8-character prefix to match.
    :return: The path of the first matching PDF file, or None if no match is found.
    """
    for file in os.listdir(folder_path):
        if file.startswith(prefix) and file.endswith('.pdf'):
            return os.path.join(folder_path, file)

    return None

def remove_invalid_filepath_characters(text):
    # List of characters that are not allowed in file names
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    
    # Replace each invalid character with a hyphen
    for char in invalid_chars:
        text = text.replace(char, '-')
    
    return text
def remove_unsupported_characters(text):
        # Implement the logic to remove unsupported characters from the text
        
        return text.encode('latin-1', 'ignore').decode('latin-1')
def main():
    # Outlook Section
    Outlook = win32com.client.Dispatch("Outlook.Application")
    print("Outlook loaded")
    # Get the public folder
    pf = Outlook.GetNamespace("MAPI").GetDefaultFolder(18).Folders.Item('Outlook Folder Name')
    print(f"Public Folder loaded: {pf.Name}")
    #Get a sub folder, in this case archive
    archive = pf.Folders.Item('Sub Folder Name')
    # Set the threshold to resume processing from where you left off
    threshold = '0'
    archiveFolderPath = f"Folder Path"
    # Loop through each project in excel, save and archive emails, save attach. for each and convert necessary files to PDF
    for inbox in archive.Folders:

        #in this case we're using project number because our email folder consists of folders starting with a 
        #8 digit project number
        if len(inbox.Name) > 8:
            projectNumber = inbox.Name[0:8]
        if projectNumber is None:
            continue
    #-------------------------------------------------------------------------------------------
    
        if projectNumber < threshold:
            print(f"Skipping project number: {projectNumber}")
            continue
        if find_pdf_with_prefix(archiveFolderPath, projectNumber):
            print(f"Skipping project number: {projectNumber} because existing PDF file was found")
            continue
    #-------------------------------------------------------------------------------------------
        print(f"missed")
        projectName = inbox.Name
        #check if project name follows scheme of four digits, hyphen, digit, digit, Letter, space, hyphen, space, project name
        #remove this if it doesn't apply
        if len(projectName) > 11:
            if projectName[4] == '-' and projectName[6].isdigit() and projectName[7].isdigit() and projectName[8].isalpha() and projectName[9] == ' ' and projectName[10] == '-' and projectName[11] == ' ':
                projectName = projectName[:8] + ' - ' + projectName[11:]
        projectName = remove_invalid_filepath_characters(projectName)
        print(f"Project Name is: {projectName}")
        
        if not os.path.exists(archiveFolderPath):
            print(f"Error: Failed to read BulkEmailArichive Folder")
        testIfFileExists = f"{archiveFolderPath}\\{projectName} - email archive.pdf"
        if os.path.exists(testIfFileExists):
            now = datetime.now()
            time_string = now.strftime("%H-%M-%S")
            testIfFileExists = f"{testIfFileExists} - {time_string}"


        # Loop through the inbox and add contents of each email to a pdf file saved at localFolderPath
        pdf = FPDF()
        pdf.set_margins(10, 10, 10)

        pageCount = 0;
        for message in inbox.Items:
            if message.Class == 43:  # Check if the item is an email message
                pdf.add_page()
                pageCount += 1
                pdf.set_font("Arial", size=12)
                pdf.cell(20, 10, txt="Page: " + str(pageCount), ln=1, align='L')
                pdf.cell(20, 10, txt="From: ", ln=0, align='L')
                pdf.multi_cell(200, 10, txt=remove_unsupported_characters(message.SenderName), align='L')
                pdf.cell(20, 10, txt="To: ", ln=0, align='L')
                pdf.multi_cell(200, 10, txt=remove_unsupported_characters(message.To), align='L')
                pdf.cell(20, 10, txt="CC: ", ln=0, align='L')
                pdf.multi_cell(200, 10, txt=remove_unsupported_characters(message.CC), align='L')
                pdf.cell(20, 10, txt="BCC: ", ln=0, align='L')
                pdf.multi_cell(200, 10, txt=remove_unsupported_characters(message.BCC), align='L')
                pdf.cell(20, 10, txt="Subject: ", ln=0, align='L')
                pdf.multi_cell(180, 10, txt=remove_unsupported_characters(message.Subject), align='L')
                pdf.ln(10)  # Add a line break
                pdf.multi_cell(200, 10, txt=remove_unsupported_characters(message.Body), align='L')
            else:
                try:
                    print(f"Item is not an email")
                    print(f"Item class = {message.Class}")
                except AttributeError:
                    print("Item is not an email and does not have a subject.")           
            message = None
        pdf.output(f"{archiveFolderPath}\\{projectName} - email archive.pdf")
        print(f"Email archive saved for project number: {projectNumber}")

try:
    main()
finally:
    # Cleanup code
    # Release COM objects
    if 'Outlook' in locals():
        Outlook = None
    if 'inbox' in locals():
        inbox = None
    if 'message' in locals():
        message = None
