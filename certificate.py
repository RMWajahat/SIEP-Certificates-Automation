import re
def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)

# call docx_replace_regex due to inputs
def replace_info(doc, name, string):
    reg = re.compile(r""+string)
    replace = r""+name
    docx_replace_regex(doc, reg , replace)

def replace_participant_name(doc, name):
    string = "{Name Surname}"
    replace_info(doc, name, string)

def replace_event_date(doc, event_date):
    string = "{EVENT DATE}"
    replace_info(doc, event_date, string)

def replace_event_name(doc, event):
    string = "{INSERT EVENT NAME}"
    replace_info(doc, event, string)

def replace_team_name(doc, name):
    string = "{team name}"
    replace_info(doc, name, string)

