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
 

def fd_recruit_converter(path_in, path_out, path_save):
    
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
    "Resides" : "value",
    "Education" : "value",
    "Profile" : "value",

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
        ],

    "Courses and Trainings" : ["Course and Training 1", "Course and Training 2", ...],
    "Key Skills" : ["Key Skill1", "Key Skill2", ...],
    "Interests" : ["Interest1", "Interest2", ...],

    }
        You must keep the following points in considration while extracting data from text:
        1. Do NOT split, rephrase or summarize list of Responsibilities. Extract each Responsibility as a complete sentence from text.
        2. Make it sure to keep the response in JSON format.
        3. If value not found then leave it empty/blank.
        4. Do not include Mobile number, Email and Home address.
        5. Summary/Personal Statement should be complete without being rephrased.

    """
    result = get_completion(test_text)
    
    
#     print("----------------------------------------------------------------")
#     print("                          Result                            ")
#     print("----------------------------------------------------------------")
#     print(result)
    
    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"|\"[Nn]ot [Mm]entioned\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))

    
#     print("----------------------------------------------------------------")
#     print("                          Dictionary                            ")
#     print("----------------------------------------------------------------")
#     print(dc)
    
    
    doc = docx.Document(formatted)


    for table in doc.tables:
            for row in table.rows:
                for i,cell in enumerate(row.cells):

    #                 doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

                    try:
                        if cell.text.strip(' :\n').lower() == 'name':
                            if dc['Name'].lower().replace(' ','') != 'value':
                                row.cells[i+1].text = dc['Name']
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'resides':
                            if dc['Resides'].lower().replace(' ','') != 'value':
                                row.cells[i+1].text = dc['Resides']
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'education':
                            if dc['Education'].lower().replace(' ','') != 'value':
                                row.cells[i+1].text = dc['Education']
                    except:
                        pass



    for i,p in enumerate(doc.paragraphs):
    
        if p.text.strip(' :\n').lower() == 'profile':
            try:
                if dc['Profile'].lower().replace(' ','') != 'value':
                    doc.paragraphs[i+2].add_run(dc['Profile'].strip())
            except:
                pass

        if p.text.strip(' :\n').lower() == 'career history':
            try:
                for j in dc['Career History']:
                    if j['Job Title'].strip() and j['Job Title'].lower().replace(' ','') !='titleofjob' or (j['Company Name'].strip() and j['Company Name'].lower().replace(' ','') !='nameofcompany'):
                        if j['Company Name'].strip(): 
                            company_run = doc.paragraphs[i+2].add_run(j['Company Name'].strip() + ' ')
                            company_run.bold = True
                        else:
                            company_run = doc.paragraphs[i+2].add_run("Company Name not mentioned")                         
                        if j['Duration'].strip():
                            duration_run = doc.paragraphs[i+2].add_run('(' + j['Duration'].strip() + ')' + '\n')
                            duration_run.bold = True
                        else:
                            duration_run = doc.paragraphs[i+2].add_run("Duration not mentioned"+'\n')                         

                        if j['Job Title'].strip():
                            job_title_run = doc.paragraphs[i+2].add_run(j['Job Title'].strip() + '\n\n')
                            job_title_run.bold = False
                        else:
                            job_title_run = doc.paragraphs[i+2].add_run("Job Title not mentioned"+'\n\n')
                            
                        if j['Responsibilities'] and j['Responsibilities'][0].lower().replace(' ','') != 'responsibility1':   
                            for k in j['Responsibilities']:
                                doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n')
                            doc.paragraphs[i+2].add_run("\n\n")

            except:
                pass



        if p.text.strip(' :\n').lower() == 'courses and trainings':
            try:
                if dc['Courses and Trainings'][0].lower().replace(' ','') != 'coursesandtrainings1':
                    for j in dc['Courses and Trainings']:
                        if j.strip():
                            language_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')

            except:
                pass

        if p.text.strip(' :\n').lower() == 'key skills':
            try:
                if dc['Key Skills'][0].lower().replace(' ','') != 'keyskill1':
                    for j in dc['Key Skills']:
                        if j.strip():
                            language_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')

            except:
                pass



        if p.text.strip(' :\n').lower() == 'interests':
            try:
                if dc['Interests'][0].lower().replace(' ','') != 'interests1':
                    for j in dc['Interests']:
                        if j.strip():
                            doc.paragraphs[i+1].add_run("\n")
                            doc.paragraphs[i+1].add_run('  • ' + j.strip())
            except:
                pass

    doc.save(path_save)
    print("Conversion completed !!")
