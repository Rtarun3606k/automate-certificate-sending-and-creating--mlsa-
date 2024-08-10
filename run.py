#!/usr/bin/env python3
import re

# Based on this link
# https://stackoverflow.com/a/42829667/11970836
# This function replace data and keeps style
def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)


        # Replace text in text boxes
def replace_ambassador2(doc_obj, regex, replace):
        for shape in doc_obj.inline_shapes:
            if shape.type == 202:  # Check if it's a text box
                textbox = shape._inline.graphic.graphicData.txbxContent
                for paragraph in textbox.p:
                    if regex.search(paragraph.text):
                        for i in range(len(paragraph.r)):
                            if regex.search(paragraph.r[i].text):
                                text = regex.sub(replace, paragraph.r[i].text)
                                paragraph.r[i].text = text




# call docx_replace_regex due to inputs
def replace_info(doc, name, string):
    reg = re.compile(r""+string)
    replace = r""+name
    docx_replace_regex(doc, reg , replace)

def replace_participant_name(doc, name):
    string = "{Name Surname}"
    replace_info(doc, name, string)

def replace_event_name(doc, event):
    string = "{INSERT EVENT NAME}"
    replace_info(doc, event, string)

def replace_ambassador_name(doc, name):
    string = "{amb}"
    replace_info(doc, name, string)

def replace2(doc,name):
    string = "{amb2}"
    replace_info(doc, name, string)


def replace_ambassador_name2(doc, name):
    student_name_regex = re.compile(r"{AMBASSADOR2}")
    replace_ambassador2(doc, name, student_name_regex)

