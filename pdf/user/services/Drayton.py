import os
import openai
import docx
import docx2txt
from .keys import api_key
from pprint import pprint
import json
import re
import textwrap
import PyPDF2
import pdfplumber
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]


# Functions to check whether the unformatted file is a docx or pdf
def read_text_from_docx(file_path):
    doc = docx.Document(file_path)
    text = [paragraph.text for paragraph in doc.paragraphs]
    return '\n'.join(text)

def read_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = []
        for page in pdf_reader.pages:
            text.append(page.extract_text())
        return '\n'.join(text)
    

def drayton_converter(path_in, path_out, path_save):
    
    formatted = path_out
    
    # unformatted document
    if path_in.endswith('.docx'):
        unformatted_text = read_text_from_docx(path_in)
    elif path_in.endswith('.pdf'):
        unformatted_text = read_text_from_pdf(path_in)
    else:
        error = 'Format not supported.'
        print(error)

    # formatted document
    formatted_text = docx2txt.process(formatted)
    
    print("Process has started...")
        
    # Prompt
    openai.api_key = api_key
    test_text = """

    Extract data from this text:

    \"""" + unformatted_text + """\"

    in following JSON format:
    {
    "Name" : "value",
    "Location" : "value",
    "Academic Qualification" : "value",
    "Personal Profile" : "value",

    "Career Experience" : [
        {"Company Name" : "Name of company",
        "Job Title" : "Title of job",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        {"Company Name" : "Name of company",
        "Job Title" : "Title of job",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        ...
        ],

    "Skills" : ["Skill1", "Skill2", ...],
    "Interest and Hobbies" : ["interest and hobbies1", "interest and hobbies2", ...],
    "Languages" : ["Language1", "Language2", ...],

    }
    
    You must keep the following points in considration while extracting data from text:
        1. Do NOT split, rephrase or summarize list of Responsibilities. Extract each Responsibility as a complete sentence from text.
        2. Make it sure to keep the response in JSON format.
        3. If value not found then leave it empty/blank.
        4. Do not include Mobile number, Email and Home address. 
    """
    result = get_completion(test_text)
    
#     print("----------------------------------------------------------------")
#     print("                          Result                            ")
#     print("----------------------------------------------------------------")
#     print(result)
    
    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))
    
#     print("----------------------------------------------------------------")
#     print("                          Dictionary                            ")
#     print("----------------------------------------------------------------")
#     print(dc)
    
    doc = docx.Document(formatted)

    for table in doc.tables:
        for row in table.rows:
            for i,cell in enumerate(row.cells):
                try:
                    if cell.text.strip(' :\n').lower() == 'name':
                        row.cells[i+1].text = dc['Name']
                except:
                    pass
                try:
                    if cell.text.strip(' :\n').lower() == 'location':
                        row.cells[i+1].text = dc['Location']
                except:
                    pass                
                try:
                    if cell.text.strip(' :\n').lower() == 'academic qualification':
                        for j in dc['Academic Qualification']:
                            row.cells[i+1].text = row.cells[i+1].text + j
                except:
                    pass                
                try:
                    if cell.text.strip(' :\n').lower() == 'personal profile':
                        personal_profile = row.cells[i+1]
                        personal_profile.text = dc['Personal Profile']

                        paragraph = personal_profile.paragraphs[0]
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                        
                except:
                    pass                

    font_size = 12

    for i,p in enumerate(doc.paragraphs):

    #         doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        if p.text.strip(' :\n').lower() == 'career experience':
            try:
                for j in dc['Career Experience']:
                    company_name = j['Company Name'].strip()
                    duration = j['Duration'].strip()
                    job_title = j['Job Title'].strip()

                    company_run = doc.paragraphs[i+2].add_run(company_name + ' ')
                    company_run.bold = True
                    company_run.font.size = Pt(font_size)

                    duration_run = doc.paragraphs[i+2].add_run('(' + duration + ')' + '\n')
                    duration_run.bold = True
                    duration_run.font.size = Pt(font_size)

                    job_title_run = doc.paragraphs[i+2].add_run(job_title + '\n\n')
                    job_title_run.bold = False
                    job_title_run.font.size = Pt(font_size)
    #                 doc.paragraphs[i+2].add_run('Duties:' + '\n\n')
                    for k in j['Responsibilities']:
                        respo = doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n')
                        respo.font.size = Pt(font_size)
                    doc.paragraphs[i+2].add_run("\n\n")
                    
            except:
                pass

        if p.text.strip(' :\n').lower() == 'skills':
            try:
                for j in dc['Skills']:
                    skill_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
                    skill_run.font.size = Pt(font_size)
            except:
                pass

        if p.text.strip(' :\n').lower() == 'interest and hobbies':
            try:
                for j in dc['Interest and Hobbies']:
                    i_h_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
                    i_h_run.font.size = Pt(font_size)
            except:
                pass

        if p.text.strip(' :\n').lower() == 'language':
            try:
                for j in dc['Languages']:
                    language_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
                    language_run.font.size = Pt(font_size)
            except:
                pass
    
    doc.save(path_save)
    print("Conversion completed !!")
    
