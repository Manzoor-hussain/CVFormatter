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
    

def scale_genesis_converter(path_in, path_out, path_save):
    
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
    "Position" : "value",
    "Availability" : "value",
    "Summary of skills" : ["Key Skill1", "Key Skill2", ...],
    "Salary Expectations/Rate" : "value",
    "Location" : "value",
    "Summary" : "value",

    "Work Experience" : [
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
    "Professional Qualifications" : ["Professional Qualification1", "Professional Qualification2", ...],
    "Languages" : ["Language1", "Language2", ...],
    "Interests" : ["Interest1", "Interest2", ...]

    }

    Please keep the following points in considration while extracting data from text:
        1. Do not summarize or rephrase Responsibilities. Extract each Responsibility completely from text.
        2. Make it sure to keep the response in JSON format.
        3. If value not found then leave it empty/blank.
        4. Do not include Mobile number, Email and Home address.
        5. Do not include Grade
    """
    result = get_completion(test_text)
    
#     print("----------------------------------------------------------------")
#     print("                          Result                            ")
#     print("----------------------------------------------------------------")
#     print(result)
    
    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Uu]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))
    
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
                            row.cells[i+1].text = dc['Name']
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'position':
                            row.cells[i+1].text = dc['Position']
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'availability':
                            row.cells[i+1].text = dc['Availability']
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'summary of skills':
                            for j in dc['Summary of skills']:
                                run = row.cells[i+1].paragraphs[0].add_run('  • ' + j.strip() + '\n')
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'salary expectations/rate':
                            for j in dc['Salary Expectations/Rate']:
                                row.cells[i+1].text = row.cells[i+1].text + j
                    except:
                        pass

                    try:
                        if cell.text.strip(' :\n').lower() == 'location':
                            row.cells[i+1].text = dc['Location']
                    except:
                        pass


    for i,p in enumerate(doc.paragraphs):

    #         doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        if p.text.strip(' :\n').lower() == 'summary':
          
            doc.paragraphs[i+2].add_run(dc['Summary'])
#             name_paragraph.runs[0].bold = True
            

        if p.text.strip(' :\n').lower() == 'work experience':
            try:
                for j in dc['Work Experience']:
                    company_name = j['Company Name'].strip()
                    duration = j['Duration'].strip()
                    job_title = j['Job Title'].strip()

                    company_run = doc.paragraphs[i+2].add_run(company_name + ' ')
                    company_run.bold = True

                    duration_run = doc.paragraphs[i+2].add_run('(' + duration + ')' + '\n')
                    duration_run.bold = True

                    job_title_run = doc.paragraphs[i+2].add_run(job_title + '\n\n')
                    job_title_run.bold = False

    #                 doc.paragraphs[i+2].add_run('Duties:' + '\n\n')
                    for k in j['Responsibilities']:
                        respo = doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n')
                    doc.paragraphs[i+2].add_run("\n\n")

            except:
                pass


        if p.text.strip(' :\n').lower() == 'education':
            try:
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
            except:
                pass


        if p.text.strip(' :\n').lower() == 'professional qualifications':
            try:
                for j in dc['Professional Qualifications']:
                    language_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')

            except:
                pass

        if p.text.strip(' :\n').lower() == 'languages':
            try:
                for j in dc['Languages']:
                    language_run = doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')

            except:
                pass



        if p.text.strip(' :\n').lower() == 'interests':
            try:
                for j in dc['Interests']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n')
            except:
                pass

    doc.save(path_save)
    print("Conversion completed !!")
    
