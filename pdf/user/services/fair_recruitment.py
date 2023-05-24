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


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, 
    )
    return response.choices[0].message["content"]


def fair_recruitment_converter(path, pathout, path_save):
    formatted = pathout
    file_path = path

    if file_path.endswith('.docx'):
        unformated_text = read_text_from_docx(file_path)
    elif file_path.endswith('.pdf'):
        unformated_text = read_text_from_pdf(file_path)
    else:
        print('Unsupported file format')



    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key
    
    print("Process has Started...")
    test_text = """

    Extract data from this text:

    \"""" + unformated_text + """\"

    in following JSON format:
    {
    "Name" : "value",
    "Summary" : "value",
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
    "Work and Team Experience" : [
        {"Company Name" : "Name of Company",
        "Duration" : "Working Duration in Company",
        "Job Title" : "Title of job",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...]
        },
        {"Company Name" : "Name of Company",
        "Duration" : "Working Duration in Company",
        "Job Title" : "Title of job",
        "Responsibilities" : ["Responsibilities1","Responsibilities2", ...]
        },
        ...
        ],
    "Achievements" : ["Achievement1","Achievement2",...],
    "Qualifications" : ["Qualification1", "Qualification2", ...],
    "Skills" : ["Skill1", "Skill2", ...],
    "Languages" : ["Language1", "Language2", ...],
    "Interests" : ["interest1", "interest2", ...]
    }
    """

    result = get_completion(test_text)

    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):


        if p.text.strip(' :\n').lower() == 'name':
            try:
                name_paragraph = doc.paragraphs[i]
                name_paragraph.text = str(dc['Name'])
                name_paragraph.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
            except:
                pass


        if p.text.strip(' :\n').lower() == 'summary':
            try:
                summary = doc.paragraphs[i] 
                summary.text = str(dc['Summary'])
                doc.paragraphs[i].add_run(summary + '\n')
                doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass


        if p.text.strip(' :\n').lower() == 'education':
            try:
                for j in dc['Education']:
                    institute_name = j['Institute Name'].strip()
                    degree_name = j['Degree Name'].strip()
                    duration = j['Duration'].strip()

                    doc.paragraphs[i+2].add_run(institute_name + '\n').bold = True
                    doc.paragraphs[i+2].add_run(degree_name + '\n').bold = True
                    doc.paragraphs[i+2].add_run(duration + '\n\n').bold = False
            except:
                pass

        if p.text.strip(' :\n').lower() == 'work and team experience':
            try:
                for j in dc['Work and Team Experience']:
                    if j['Company Name']:
                        doc.paragraphs[i+2].add_run(j['Company Name'].strip() + '\n').bold = True
                    if j['Duration']:
                        doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n').bold = True
                    if j['Job Title']:
                        doc.paragraphs[i+2].add_run(j['Job Title'].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run('\n')
                    for k in j['Responsibilities']:
                        doc.paragraphs[i+2].add_run('     •   '+ k.strip() + '\n').bold = False
                    doc.paragraphs[i+2].add_run('\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'achievements':
            try:
                for j in dc['Achievements']:
                    doc.paragraphs[i+2].add_run('     •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'qualifications':
            try:
                for j in dc['Qualifications']:
                    doc.paragraphs[i+2].add_run('     •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'skills':
            try:
                for j in dc['Skills']:
                    doc.paragraphs[i+2].add_run('     •   ' + j.strip() + '\n')
                    doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass

        if p.text.strip(' :\n').lower() == 'languages':
            try:
                for j in dc['Languages']:
                    doc.paragraphs[i+2].add_run('     •   '  + j.strip() + '\n')
            except:
                pass


        if p.text.strip(' :\n').lower() == 'interests':
            try:
                for j in dc['Interests']:
                    doc.paragraphs[i+2].add_run('     •   ' + j.strip() + '\n')
                    doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass


    doc.save(path_save)
    print("Procees has Completed...")
