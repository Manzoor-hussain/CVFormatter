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
    
    
def harrington_morris_converter(path_in, path_out, path_save):
    
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
    
    
    openai.api_key = api_key
    test_text = """

    Extract data from this text:

    \"""" + unformatted_text + """\"

    in following JSON format:
    {
    "Name" : "value",
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
    "Software Skills" : ["Software Skill1", "Software Skill2", ...],
    "Certifications" : ["Certification1", "Certification2", ...],
    "Professional Qualification" : "value",
    "Languages" : "value",
    "Nationality" : "value",

    "Professional Experience" : [
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

    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))
    
#     print("----------------------------------------------------------------")
#     print("                          Dictionary                            ")
#     print("----------------------------------------------------------------")
#     print(dc)

    
    doc = docx.Document(formatted)

    for table in doc.tables:
        for row in table.rows:
            for i,cell in enumerate(row.cells):
                
#                 doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                
            
                if cell.text.strip(' :\n').lower() == 'education':
                    for j in dc['Education']:
                        institute_name = j['Institute Name'].strip()
                        duration = j['Duration'].strip()
                        degree_name = j['Degree Name'].strip()

                        run = row.cells[i+1].paragraphs[0].add_run(institute_name + ' ')
                        run.bold = True

                        if duration:
                            run = row.cells[i+1].paragraphs[0].add_run('(' + duration + ')' + '\n')
                            run.bold = True
                        else:
                            run = row.cells[i+1].paragraphs[0].add_run('(' + "Not mentioned" + ')' + '\n')
                            run.bold = True

                        if degree_name:
                            run = row.cells[i+1].paragraphs[0].add_run(degree_name + '\n\n')
                            run.bold = False
                        else:
                            run = row.cells[i+1].paragraphs[0].add_run("Not mentioned" + '\n\n')
                            run.bold = False
                
                
                try:
                    if cell.text.strip(' :\n').lower() == 'software skills':
                        for j in dc['Software Skills']:
                            row.cells[i+1].paragraphs[0].add_run('  • ' + j.strip() + '\n')
                except:
                    pass
                
                try:
                    if cell.text.strip(' :\n').lower() == 'certifications':
                        for j in dc['Certifications']:
                            row.cells[i+1].paragraphs[0].add_run('  • ' + j.strip() + '\n')
                except:
                    pass
                
                try:
                    if cell.text.strip(' :\n').lower() == 'professional qualification':
                        row.cells[i+1].text = dc['Professional Qualification']
                except:
                    pass
                
                try:
                    if cell.text.strip(' :\n').lower() == 'languages':
                        for j in dc['Languages']:
                            row.cells[i+1].paragraphs[0].add_run('  • ' + j.strip() + '\n')
                except:
                    pass
                
                try:
                    if cell.text.strip(' :\n').lower() == 'nationality':
                        row.cells[i+1].text = dc['Nationality']
                except:
                    pass

    
    font_size = 14
    for i,p in enumerate(doc.paragraphs):

    #         doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        
        if p.text.strip(' :\n').lower() == 'name':
            try:
                name_paragraph = doc.paragraphs[i]
                name_paragraph.text = str(dc['Name'] + '\n')
                name_paragraph.runs[0].bold = True
                name_paragraph.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
                name_paragraph.runs[0].font.size = Pt(font_size)
            except:
                pass
            
            
        if p.text.strip(' :\n').lower() == 'professional experience':

                for j in dc['Professional Experience']:
                    
                    doc.paragraphs[i+1].add_run("\n")
                    company_name = j['Company Name'].strip()
                    duration = j['Duration'].strip()
                    job_title = j['Job Title'].strip()

                    company_run = doc.paragraphs[i+1].add_run(company_name + ' ')
                    company_run.bold = True

                    duration_run = doc.paragraphs[i+1].add_run('(' + duration + ')' + '\n')
                    duration_run.bold = True

                    job_title_run = doc.paragraphs[i+1].add_run(job_title + '\n\n')
                    job_title_run.bold = False

    #                 doc.paragraphs[i+2].add_run('Duties:' + '\n\n')
                    for k in j['Responsibilities']:
                        respo = doc.paragraphs[i+1].add_run('  • ' + k.strip() + '\n')
                    doc.paragraphs[i+1].add_run("\n\n")    
        

    doc.save(path_save)
    print("Coversion Completed !!!")

    
