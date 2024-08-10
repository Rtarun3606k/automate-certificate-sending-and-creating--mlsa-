#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
from run import *
from docx import Document
import csv
from docx2pdf import convert
from sendmail import send_email


# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


def get_participants(f):
    data = [] # create empty list
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) # append all results
        print(data)
    return data

def create_docx_files(filename, list_participate, ambassador,eventname):
    print(list_participate,"list particip[ante]")
    for participate in list_participate:
        # use original file everytime
        print(participate["Name"],participate['email'])
        doc = Document(filename)

        capitalized_text_not = participate["Name"]
        # text = "your example text here"
        name = ' '.join(word.capitalize() for word in capitalized_text_not.split())
        event = eventname

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)
        replace2(doc,ambassador)
        doc.save('Output/Doc/{}.docx'.format(name))

        # ! if your program working slowly, comment this one line and open other 2 line.
        print("Output/{}.pdf Creating".format(name))
        convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))
        body = "body"
        # send_email(participate['Name'],participate['email'], eventname, 'Output/Pdf/{}.pdf'.format(name))

        # ! Open those lines and comment above 2 lines if your program working extremely slow
        # os.system("docx2pdf Output/Doc/")
        # os.system("move Output\Doc\*pdf Output\PDF")
      

    
# get certificate temple path
certificate_file = "Event Certificate Template2.docx"
# get participants path
participate_file = "test.csv"

# Enter your name here [Ambassador Name]
# ambassador_name = input("Enter ambassador name: ")
ambassador_name = "Tarun Nayaka R"


# get participants
list_participate = get_participants(participate_file);
# eventname = input("Enter event Name: ")
eventname = "Hosting a Web Application on Azure - Day 1"

# process data
create_docx_files(certificate_file, list_participate, ambassador_name,eventname)

 

 
 
 
 




