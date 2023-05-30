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
    )
    return response.choices[0].message["content"]


def cw_executive_converter(path, pathout, path_save):
    formatted= pathout
    file_path = path  

    if file_path.endswith('.docx'):
        unformated_text = read_text_from_docx(file_path)
    elif file_path.endswith('.pdf'):
        unformated_text = read_text_from_pdf(file_path)
    else:
        print('Unsupported file format')


    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key

    print ("Process has Started...")
    test_text = """

    Extract data from this text:

    \"""" + unformated_text + """\"

    in following JSON format:
    {
    "Name" : "value",
    "Professional Profile" : "value",
    "Career History" : [
        {"Duration" : "Working Duration in Company",
        "Job Title" : "Title of job",
        "Company Name" : "Name of Company",
        "Responsibilities" : ["Responsiblility1, Responsibility2", ...]
        },
        {"Duration" : "Working Duration in Company",
        "Job Title" : "Title of job",
        "Company Name" : "Name of Company",
        "Responsibilities" : ["Responsiblility1, Responsibility2", ...]
        },
        ...
        ],
    "Education" : [
        {"Duration" : "Studying duration in institute",
        "Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree"
        },
        {"Duration" : "Studying duration in institute",
        "Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree"
        },
        ...
        ],
    "Achievements" : ["Achievement1","Achievement2",...],
    "Qualifications" : ["Qualification1", "Qualification2", ...],
    "Skills" : ["Skill1", "Skill2", ...],
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




    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))
    dc





    doc = docx.Document(formatted)


    def change_font_size(paragraph, size):
        for run in paragraph.runs:
            run.font.size = size

    for i, p in enumerate(doc.paragraphs):
        if p.text.strip(' :\n').lower() == 'name':
            try:
                name = doc.paragraphs[i]
                name.text = dc['Name']
                name.alignment = docx.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in name.runs:
                    run.bold = True
                    run.font.size = docx.shared.Pt(16) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'professional profile':
            try:
                summary = doc.paragraphs[i+1] 
                summary.text = str(dc['Professional Profile'])
                doc.paragraphs[i+1].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                change_font_size(summary, docx.shared.Pt(12)) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'career history':
            try:
                for j in dc['Career History']:
                    if j['Duration']:
                        doc.paragraphs[i+1].add_run(j['Duration'].strip() + '\n').bold = True
                    if j['Job Title']:
                        doc.paragraphs[i+1].add_run(j['Job Title'].strip() + '\n').bold = True
                    if j['Company Name']:
                        doc.paragraphs[i+1].add_run(j['Company Name'].strip() + '\n\n').bold = True
                    if j['Responsibilities']:
                        for k in j['Responsibilities']:
                            doc.paragraphs[i+1].add_run('•   ' + k.strip() + '\n').bold = False
                    doc.paragraphs[i+1].add_run('\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'education':
            try:
                for j in dc['Education']:
                    if j['Duration']:
                        duration = j['Duration'].strip()
                        doc.paragraphs[i+1].add_run(duration + '\n').bold = True
                    if j['Degree Name']:
                        degree_name = j['Degree Name'].strip()
                        doc.paragraphs[i+1].add_run(degree_name + '\n').bold = True
                    if j['Institute Name']:
                        institute_name = j['Institute Name'].strip()
                        doc.paragraphs[i+1].add_run(institute_name + '\n').bold = True
                        doc.paragraphs[i+1].add_run('\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12))
            except:
                pass

        if p.text.strip(' :\n').lower() == 'achievements':
            try:
                for j in dc['Achievements']:
                    doc.paragraphs[i+1].add_run('•   ' + j.strip() + '\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12)) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'qualifications':
            try:
                for j in dc['Qualifications']:
                    doc.paragraphs[i+1].add_run('•   ' + j.strip() + '\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12)) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'skills':
            try:
                for j in dc['Skills']:
                    doc.paragraphs[i+1].add_run('•   ' + j.strip() + '\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12)) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'languages':
            try:
                for j in dc['Languages']:
                    doc.paragraphs[i+1].add_run('•   ' + j.strip() + '\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12)) 
            except:
                pass

        if p.text.strip(' :\n').lower() == 'interests':
            try:
                for j in dc['Interests']:
                    doc.paragraphs[i+1].add_run('•   ' + j.strip() + '\n')
                    change_font_size(doc.paragraphs[i+1], docx.shared.Pt(12)) 
            except:
                pass


    doc.save(path_save)
    print("Process has Completed...")

