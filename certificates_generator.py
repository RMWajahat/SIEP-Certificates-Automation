# Comptech UET Taxila Campus holds rights to this script @ 2022
import os
from certificate import *
from docx import Document
import csv
from docx2pdf import convert


# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


def get_participants(f):
    data = [] 
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) 
    return data

def create_docx_files(filename, list_participate, manager_team):

    for participate in list_participate:
        doc = Document(filename)

        name = participate["Name Surname"].strip()
        event = participate["Event"].strip()
        event_date = participate["Event Date"]

        if name == "" or event == "" or event_date == "":
            continue
        else:
            replace_participant_name(doc, name)
            replace_event_name(doc, event)
            replace_event_date(doc, event_date)
            # replace_team_name(doc, manager_team)      # you can use in case of any event collaboration
            doc.save('Output/Doc/{}.docx'.format(name))

            print("Output/{}.pdf Creating".format(name))
            convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))

    
# get certificate temple path
certificate_file = "./Certificate_Template/sample.docx"
# get participants path
participants_list = "./Participants/Participants.csv"

# get participants
list_participate = get_participants(participants_list)

# process data
create_docx_files(certificate_file, list_participate, "`")



