#########################################################
#########################################################
######                                             ######
######     MSG to CSV for ALJR to LEX BCRO         ######
######                                             ######
######     This file starts with a folder          ######
######     of Email files (.msg), extract_msg      ######
######     infromation from the emails and         ######
######     saves the PDF files locally. Then       ######
######     it extaracts additioanl info            ######
######     from the PDFs and saves                 ######
######     everyuthing to a .CSV file              ######
######                                             ######
######     Abdulla Arid                            ######
######     Dylan Hunter                            ######
######     Alexander Joubert                       ######
######                                             ######
######     Last edit: April 18, 2023               ######
######                                             ######
#########################################################
#########################################################

#################
### LIBRARIES ###
#################

import os
import re
import csv
import uuid
import PyPDF2
import datetime
import unicodedata
import extract_msg
from dateutil.parser import parse

#################
### FUNCTIONS ###
#################

def normalize_text(text):
    """
    Used to normalize the text scraped from PDF files
    """
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8', 'ignore')
# end normalize_text

def generate_unique_id():
    """
    Generates a unique ID for every RPA record generated
    """
    return str(uuid.uuid4())
# end generate_unique_id

def seperate_first_last(full_name):
    """
    Takes in a full name and outputs first and last names
    """
    if not full_name: # error check for Null input
        return None
    name_parts = full_name.split()
    if len(name_parts) > 1:
        first_name = ' '.join(name_parts[:-1])
        last_name = name_parts[-1]
    else:
        first_name = full_name
        last_name = ''
    return first_name, last_name
# end seperate_first_last

def check_minister(input_list):
    """
    Takes in the scraped Minister titles and outputs the title desired by client
    """
    if not input_list:  # error check for Null input
        return None
    first_item = input_list[0]
    # makes all things caps for easier processing
    if not first_item.isupper():
        first_item = first_item.upper()
    # Minsister check 1 english
    if "SAFETY" in first_item or "EMERGENCY" in first_item:
        return ['Minister of Public Safety']
    # Minsister check 2 english
    if "CITIZENSHIP" in first_item or "IMMIGRATION" in first_item or "REFUGEES" in first_item:
        return ['Minister of Immigration, Refugees and Citizenship']
    # Minsister check 3 french
    if "L’IMMIGRATION" in first_item or "RÉFUGIÉS" in first_item or "CITOYENNETÉ" in first_item:
        return ['Ministre de l’Immigration, des Réfugiés et de la Citoyenneté']
    # Minsister check 4 french
    if "SÉCURITÉ" in first_item or "PUBLIQUE" in first_item:
        return ['Ministre de la Sécurité publique']
    # if no matching ministers returns the original input
    else:
        return input_list
# end check_minister

def normalize_numbers(input_list):
    """
    Takes a list of dirty numerical input and outputs it as clean numbers only
    """
    if not input_list:  # error check for Null input
        return None
    result = []
    for input_str in input_list:
        numbers = ''
        for char in input_str:
            if char.isdigit():
                numbers += char
        result.append(numbers.rstrip())
        if result and result[-1].endswith('\n'):
            result[-1] = result[-1][:-1]
    return result if result else None
# end normalize_numbers

########################
### FILE DIRECTORIES ###
########################

msg_folder_path = 'C:\\Python\\Project 1 - ALJR to LEX\\Samples\\Emails'
pdf_output_folder = 'C:\\Python\\Project 1 - ALJR to LEX\\Extracted PDFs'
timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
csv_file_path = f'C:\\Python\\Project 1 - ALJR to LEX\\CSV Output\\ALJR to LEX {timestamp}.csv'

###########################
### REGULAR EXPRESSIONS ###
###########################

Study_Permit = re.compile(r'[sS]tudy [pP]ermit')
Court_File_Number = re.compile(r'(?s)\bIMM-\d{3,5}-\d{2}')
Applicant_Num = re.compile(r'applicantnull')
Uci_Num = re.compile(r'ucinull')
Application_Num = re.compile(r'applicationnull')
Applicant_Name = re.compile(r'(?is)\bBetween:\s*(.*?\s)\bApplicant\b')
First_Name = re.compile('')
Last_Name = re.compile('')
Respondent = re.compile(r'^(?:The\s)?Minister.*$', re.IGNORECASE | re.MULTILINE)
UCI_Number = re.compile(r'\b(?:\d{10}|(?:\d{2}\s*-\s*\d{4}\s*-\s*\d{4})|\d{8})(?:(?<=\d)\b|(?<=\d)\.)', re.MULTILINE)
Application_Number = re.compile(r'(?s)(?<=\s)[A-Za-z]\s?\d{9,10}')
Address = re.compile(r'(?:.*\n){4}.*[A-Za-z]\d[A-Za-z] \d[A-Za-z]\d')
Address_Count = re.compile('')
Study_Permit = re.compile(r'[sS]tudy [pP]ermit')
Date = re.compile(r'.*[Dd][Aa][Tt][Ee][Dd] at ([^\n\r]*)')
Date_Trim = re.compile(r'\b([A-Za-z]+\s+\d{1,2},\s+\d{4})\b')
file_number_regex = re.compile(r'numéro de dossier de la cour', re.IGNORECASE)
fed_regex = re.compile(r'cour[ \-]?fédérale', re.IGNORECASE)
review_regex = re.compile(r'Demande d’autorisation et demande de contrôle', re.IGNORECASE)
Prepared_By = re.compile(r'.*?[pP]repared by (.*?)(?=\.|:|$)')
# pdf_text_regex = re.compile(r'')

#######################
### DATA DICTIONARY ###
#######################

regexDict = {
    'RPA PDF Study Permit': Study_Permit,
    'RPA PDF Court File Number': Court_File_Number,
    'RPA PDF Number of Applicants': Applicant_Num,
    'RPA PDF Applicant Name': Applicant_Name,
    'RPA PDF First Name': First_Name,
    'RPA PDF Last Name': Last_Name,
    'RPA PDF Respondent': Respondent,
    'RPA PDF UCI Number': UCI_Number,
    'RPA PDF Number of UCI Numbers': Uci_Num,
    'RPA PDF Application Number': Application_Number,
    'RPA PDF Number of Application Numbers': Application_Num,
    'RPA PDF Date': Date,
    'RPA PDF Prepared By': Prepared_By,
    'RPA PDF Address': Address,
    'RPA PDF Address Count': Address_Count,
    # 'RPA PDF Text': pdf_text_regex,
}

#################
### VARIABLES ###
#################

pdf_count = 0
file_count = 0
processed_files = 0

#################
### MAIN CODE ###
#################

# Counts the number of files in the selected folder
for _, _, files in os.walk(msg_folder_path):
    for file in files:
        if file.lower().endswith('.msg'):
            file_count += 1

# Open the CSV file for writing
with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:

    # Create a CSV writer
    csv_writer = csv.writer(csv_file,
                            delimiter=',',
                            quotechar='"',
                            quoting=csv.QUOTE_MINIMAL,
                            escapechar='\\')
    
    # Write the header row
    csv_writer.writerow(['RPA Record Number',           # A : RPA Record Number
                         'RPA Record Status',           # B : RPA Record Status
                         'RPA Record Created On',       # C : RPA Record Created On
                         'RPA Record Region',           # D : RPA Record Region
                         'RPA Record Type',             # E : RPA Record Type
                         'RPA Email File Name',         # F : RPA Email File Name
                         'RPA Email Attachment Name',   # G : RPA Email Attachment Name
                         'RPA Email Attachments Count', # H : RPA Email Attachments Count
                         'RPA Email Sender',            # I : RPA Email Sender
                         'RPA Email Receiver',          # J : RPA Email Receiver
                         'RPA Email Subject',           # K : RPA Email Subject
                         'RPA Email Date',              # L : RPA Email Date
                         'RPA Email Body',              # M : RPA Email Body
                         'RPA PDF Language']            # N : RPA PDF Language
                        + list(regexDict.keys()))       # O <--> AA

    # Loop through all the files in the folder
    for file_name in os.listdir(msg_folder_path):

        # Check if the file is a .msg file
        if file_name.endswith('.msg'):
            msg_file_path = os.path.join(msg_folder_path, file_name)

            # Open the .msg file
            msg = extract_msg.Message(msg_file_path)

            # Get email information
            sender = msg.sender
            receiver = '; '.join([recipient.email if hasattr(recipient, 'email') else recipient for recipient in msg.to]).replace('; ', '').replace(' ', '')
            subject = msg.subject
            date = msg.date
            body = msg.body

            # Initialize PDF attachment counter
            pdf_attachments_count = 0

            # Loop through the attachments
            for attachment in msg.attachments:

                # Check if the attachment is a PDF by file extension and MIME type
                if (attachment.longFilename.lower().endswith('.pdf') and attachment.mimetype.lower() == 'application/pdf'):
                    
                    # Save the attachment to the output folder
                    attachment_name = attachment.save(customPath=pdf_output_folder)

                    # Increment the PDF attachment counter
                    pdf_attachments_count += 1

                    # add the email and attachment information to a list
                    line = [file_name,                  # F : Email Name
                            attachment.longFilename,    # G : Attachment Name
                            pdf_attachments_count,      # Q : PDF Attachment Count
                            sender,                     # I : Sender
                            receiver,                   # J : Reciever
                            subject,                    # K : Subject
                            date,                       # L : Date
                            body]                       # M : Body

                    # attachment handeling 
                    pdf_file_path = os.path.join(pdf_output_folder, attachment_name)
                    pdf_file = open(pdf_file_path, 'rb')
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    text = ''
                    page_count = len(pdf_reader.pages)

                    # Loop through all the pages in the PDF
                    for i in range(page_count):
                        page = pdf_reader.pages[i]
                        text += page.extract_text()
                        text = normalize_text(text)

                    # N : RPA PDF Language
                    if file_number_regex.search(text) and fed_regex.search(text):
                        language = 'French'
                    else:
                        language = 'English'

                    ##############################
                    ### REGEX LOOP STARTS HERE ###
                    ##############################

                    matches = []

                    for key in regexDict:

                        regex = regexDict.get(key)
                        match = regex.findall(text)

                        # O : Study Permit
                        if (regex.pattern == Study_Permit.pattern):
                            match = Study_Permit.findall(text)
                            # print('Got to Study Permit')

                        # P : Court File Number
                        if (regex.pattern == Court_File_Number.pattern):
                            match = Court_File_Number.findall(text)
                            # print('Got to Court File Number')

                        # Q : RPA PDF Number of Applicants
                        if (regex.pattern == Applicant_Num.pattern): # key is applicant name
                            match = Applicant_Name.findall(text)
                            if (len(match) != 0): # if there are more than 0 matches
                                # solit the text by ',', 'and' or '\n'
                                nameSet = re.split(re.compile(r',|[aA][nN][dD]|\n'), match[0])
                                # remove trailing and following whitespace
                                nameSet = [x.strip(' ') for x in nameSet]
                                # add to a new list
                                new_list = [x for x in nameSet if x != '']
                                # prints the number of applicants
                                match = ((f'{len(new_list)}')) 
                            else:
                                match = str(0)
                
                        # R : RPA PDF Applicant Name
                        # S : RPA PDF First Name
                        # T : RPA PDF Last Name
                        if (regex.pattern == Applicant_Name.pattern):
                            match = Applicant_Name.findall(text)
                            if match:
                                full_name = match[0].strip()
                                first_name, last_name = seperate_first_last(full_name)
                            else:
                                full_name, first_name, last_name = '', '', ''                         
                            matches.extend([full_name, first_name, last_name])

                        # U : RPA PDF Respondent
                        if (regex.pattern == Respondent.pattern):
                            match = Respondent.findall(text)
                            print(f'Got to Respondent: {match}')
                            match = check_minister(match)
                            print(f'Got to Respondent: {match}')

                        # V : RPA PDF UCI Number
                        if (regex.pattern == UCI_Number.pattern):
                            match = UCI_Number.findall(text)
                            print(f'UCI before: {match}')
                            match = normalize_numbers(match)
                            print(f'UCI after: {match}')

                        # W : RPA PDF Number of UCI Numbers
                        if (regex.pattern == Uci_Num.pattern):
                            match = UCI_Number.findall(text)
                            # print('Got to Number of UCI Numbers')
                            if (len(match) != 0):
                                match = str(len(match))
                            else:
                                match = str(0)
                        
                        # X : RPA PDF Application Number
                        if (regex.pattern == Application_Num.pattern):
                            match = Application_Number.findall(text)

                        # Y : RPA PDF Number of Application Numbers

                        # Z : RPA PDF Date                        
                        if (regex.pattern == Date.pattern):
                            initial = re.search(Date, text)
                            if initial:
                                initial_group = initial.group(1)
                                match = re.findall(Date_Trim, initial_group)

                        # AA : RPA PDF Prepared By

                        # AB : RPA PDF Address

                        # AC : RPA PDF Address Count
                        if match is not None and len(match) > 0:
                            if key == 'RPA PDF Address':
                                count = len(match)
                                count_addresses = count
                                print(f'count_addresses: {count}')
                                text2 = ''
                                for i in match:
                                    text2 += i + '\n\n'
                                matches.append(text2)
                            elif key == 'RPA PDF Address Count':
                                matches.append(count_addresses)
                                print('got to Address Count')
                                print(f'Address count is {count_addresses}')
                            elif key == 'RPA PDF UCI Number':
                                text2 =''
                                for i in match:
                                    text2 += i + '\n'
                                matches.append(text2)
                            elif key == 'RPA PDF Application Number':
                                text2 =''
                                for i in match:
                                    text2 += i + '\n'
                                matches.append(text2)
                            elif key == 'RPA PDF Number of Application Numbers':
                                matches.append(str(len(match)))
                            elif key == 'RPA PDF First Name':
                                continue
                            elif key == 'RPA PDF Last Name':
                                continue
                            elif key == 'RPA PDF Applicant Name':
                                continue
                            elif key == 'PDF Text':
                                pdf_text=pdf_text.replace('\u25a0', '')
                                pdf_text=pdf_text.replace('\u25c4', '')
                                matches.append(pdf_text)
                            else: # use the first item in the list
                                matches.append(match[0])
                        else:
                            if (key != 'RPA PDF Applicant Name'):
                                matches.append('')
                                print('got to RPA PDF Applicant Name')

                    # Creates a Unique ID for each transaction
                    unique_id = generate_unique_id()

                    # Writes the data to the CSV file in order
                    line.insert(0, unique_id)
                    line.insert(1, "Record Status")
                    line.insert(2, timestamp)
                    line.insert(3, "BRCO")
                    line.insert(4, "ALJR")
                    line.append(language)
                    line += matches
                    csv_writer.writerow(line)

            # Close the .msg file
            msg.close()

        # Increment the overall counter
        processed_files += 1

        # Print the current file being processed / finished
        print(f"Processed file {processed_files} of {file_count} files")
        print("****************************************************************************************************************")
    
    # Print that the processes is finished
    print(f"\nFinished processing {processed_files} .msg files\n")

# Main fin