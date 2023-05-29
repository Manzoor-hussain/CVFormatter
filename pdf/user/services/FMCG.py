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



def fmcg_converter(path_in, path_out, path_save):
    
    formatted = path_out 
    
    if path_in.endswith('.docx'):
        unformatted_text = read_text_from_docx(path_in)
    elif path_in.endswith('.pdf'):
        unformatted_text = read_text_from_pdf(path_in)
    else:
        error = 'Format not supported.'
        print(error)
    
    formatted_text = docx2txt.process(formatted)
    
    
    print("Process has started...")
    
    # Prompt
    openai.api_key = api_key

    test_text = """

    Extract data from this text:

    \"""" + unformatted_text + """\"

    in following JSON format:
    {
    "Current Employer" : "value",
    "Job title" : "value",
    "Location" : "value",
    "Salary Sought" : "value",
    "Notice Period" : "value",

    "Name" : "value",
    "Profile" : "value",
    "Education" : [
        {"Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute",
        },
        {"Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute",
        },
        ...
        ],
    "Professional Qualifications" : ["Qualification1", "Qualification2", ...],
    "Skills" : ["Skill1", "Skill2", ...],
    "IT Skills" : ["IT Skill1", "IT Skill2", ...],
    "Activities" : ["Activity1", "Activity2", ...],
    "Interests" : ["interest1", "interest2", ...],
    "Languages" : ["Language1", "Language2", ...],
    "Career History" : [
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
        ]
    }
    make it sure to keep the response in JSON format.
    If value not found then leave it empty/blank.
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
                    if cell.text.strip(' :\n').lower() == 'current employer':
                        row.cells[i+1].text = dc['Current Employer']
                except:
                    pass
                try:
                    if cell.text.strip(' :\n').lower() == 'job title':
                        row.cells[i+1].text = dc['Job Title']
                except:
                    pass                
                try:
                    if cell.text.strip(' :\n').lower() == 'location':
                        for j in dc['Location']:
                            row.cells[i+1].text = row.cells[i+1].text + j
                except:
                    pass                
                try:
                    if cell.text.strip(' :\n').lower() == 'salary sought':
                        row.cells[i+1].text = dc['Salary Sought']
                except:
                    pass                
                try:
                    if cell.text.strip(' :\n').lower() == 'notice period':
                        row.cells[i+1].text = dc['Notice Period']
                except:
                    pass                



    for i,p in enumerate(doc.paragraphs):

#         doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        if p.text.strip(' :\n').lower() == 'name':
            try:
                name_paragraph = doc.paragraphs[i]
                name_paragraph.text = str(dc['Name'])
                name_paragraph.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
                name_paragraph.runs[0].bold = True
            except:
                pass


        if p.text.strip(' :\n').lower() == 'profile':
            try:
                doc.paragraphs[i+2].text = str(dc['Profile'])
                doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass


        if p.text.strip(' :\n').lower() == 'education':
            
            for j in dc['Education']:
                institute_name = j['Institute Name'].strip()
                duration = j['Duration'].strip()
                degree_name = j['Degree Name'].strip()

                doc.paragraphs[i+2].add_run(institute_name + ' ').bold = True
                if duration:
                    doc.paragraphs[i+2].add_run('(' + duration + ')' + '\n').bold = True
                else:
                    doc.paragraphs[i+2].add_run('(' + "Not mentioned" + ')' + '\n').bold = True
                if degree_name:
                    doc.paragraphs[i+2].add_run(degree_name + '\n\n').bold = False
                else:
                    doc.paragraphs[i+2].add_run("Not mentioned" + '\n\n').bold = False 
            

        if p.text.strip(' :\n').lower() == 'professional qualifications':
            try:
                for j in dc['Professional Qualifications']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass


        if p.text.strip(' :\n').lower() == 'skills':
            try:
                for j in dc['Skills']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass


        if p.text.strip(' :\n').lower() == 'it skills':
            try:
                for j in dc['IT Skills']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass        



        if p.text.strip(' :\n').lower() == 'activities':
            try:
                for j in dc['Activities']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'interests':
            try:
                for j in dc['Interests']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'languages':
            try:
                for j in dc['Languages']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'career history':
            try:
                for j in dc['Career History']:
                    company_name = j['Company Name'].strip()
                    duration = j['Duration'].strip()
                    job_title = j['Job Title'].strip()

                    doc.paragraphs[i+2].add_run(company_name + ' ').bold = True
                    doc.paragraphs[i+2].add_run('(' + duration + ')' + '\n').bold = True
                    doc.paragraphs[i+2].add_run(job_title + '\n\n').bold = False
    #                 doc.paragraphs[i+2].add_run('Duties:' + '\n\n')
                    for k in j['Responsibilities']:
                        doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n')
                    doc.paragraphs[i+2].add_run("\n\n")
#                     doc.paragraphs[i+2].add_run('\n')
            except:
                pass


    doc.save(path_save)
    print("Conversion has completed !!")