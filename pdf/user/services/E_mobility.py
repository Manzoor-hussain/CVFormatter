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


def e_mobility_converter(path, pathout, path_save):
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
    "Current Company" : "value",
    "Position applied" : "value",
    "Location" : "value",
    "Notice period" : "value",
    "Reason for Leaving" : "value",
    "System Used" : "value",
    "Dealbreakers" : "value",
    "Candidate Summary" : "value",
    "Experience" : [
        {"Job Title" : "Title of job",
        "Company Name" : "Name of Company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...]
        },
        {"Job Title" : "Title of job",
        "Company Name" : "Name of Company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...]
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
    "Publications" : ["Publication1","Publication2",...],
    "Projects" : ["Project1","Project2",...],
    "Qualifications" : ["Qualification1", "Qualification2", ...],
    "Certifications" : ["Certification1","Certification2",...],
    "Achievements" : ["Achievement1","Achievement2",...],
    "Skills" : ["Skill1", "Skill2", ...],
    "Languages" : ["Language1", "Language2", ...],
    "Interests" : ["interest1", "interest2", ...]
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
                    if cell.text.strip(' :\n').lower() == 'current company':
                        row.cells[i+1].text = dc['Current Company']
                except:
                    pass

                try:
                    if cell.text.strip(' :\n').lower() == 'position applied':
                        row.cells[i+1].text = dc['Position applied']
                except:
                    pass              

                try:
                    if cell.text.strip(' :\n').lower() == 'location':
                        row.cells[i+1].text = dc['Location']
                except:
                    pass                  

                try:
                    if cell.text.strip(' :\n').lower() == 'notice period':
                        row.cells[i+1].text = ['Notice period']
                except:
                    pass                

                try:
                    if cell.text.strip(' :\n').lower() == 'reason for leaving':
                        row.cells[i+1].text = dc['Reason for Leaving']
                except:
                    pass                

                try:
                    if cell.text.strip(' :\n').lower() == 'system used':
                        row.cells[i+1].text = dc['System Used']
                except:
                    pass

                try:
                    if cell.text.strip(' :\n').lower() == 'dealbreakers':
                        row.cells[i+1].text = dc['Dealbreakers']
                except:
                    pass

    for i,p in enumerate(doc.paragraphs):

        if p.text.strip(' :\n').lower() == 'candidate summary':
            try:
                if dc['Candidate Summary']:
                    summary = doc.paragraphs[i+1] 
                    summary.text = str(dc['Candidate Summary'])
                    doc.paragraphs[i+1].add_run('    •   ' + summary)
    #                 doc.paragraphs[i+1].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass

        if p.text.strip(' :\n').lower() == 'experience':
            try:
                for j in dc['Experience']:
                    if j['Job Title']:
                        job_title = j['Job Title'].strip()
                        doc.paragraphs[i+1].add_run(job_title + '\n').bold = True
                    if j['Company Name']:
                        company_name = j['Company Name'].strip()
                        doc.paragraphs[i+1].add_run(company_name + '\n').bold = True
                    if j['Duration']:
                        duration = j['Duration'].strip()
                        doc.paragraphs[i+1].add_run(duration + '\n').bold = True
                    doc.paragraphs[i+1].add_run('\n')
                    reponsibility = j['Responsibilities']
                    print()
                    if reponsibility:
                        doc.paragraphs[i+1].add_run('Responsibilities\n').bold = True
                        for r in reponsibility:
                            doc.paragraphs[i+1].add_run('    •   ' + r + '\n')
    #                         doc.paragraphs[i+1].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                        doc.paragraphs[i+1].add_run('\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'education':
            try:
                for j in dc['Education']:
                    if j['Institute Name']:
                        institute_name = j['Institute Name'].strip()
                        doc.paragraphs[i+1].add_run(institute_name + '\n').bold = False
                    if j['Degree Name']:
                        degree_name = j['Degree Name'].strip()
                        doc.paragraphs[i+1].add_run(degree_name + '\n').bold = False
                    if j['Duration'].strip():
                        duration = j['Duration'].strip()
                        doc.paragraphs[i+1].add_run(duration + '\n').bold = False
                    doc.paragraphs[i+1].add_run("\n")
            except:
                pass

        if p.text.strip(' :\n').lower() == 'publications':
            try:
                for j in dc['Publications']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'projects':
            try:
                for j in dc['Projects']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'qualifications':
            try:
                for j in dc['Qualifications']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'certifications':
            try:
                for j in dc['Certifications']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'achievements':
            try:
                for j in dc['Achievements']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass

        if p.text.strip(' :\n').lower() == 'skills':
            try:
                for j in dc['Skills']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
    #                 doc.paragraphs[i+1].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass

        if p.text.strip(' :\n').lower() == 'languages':
            try:
                for j in dc['Languages']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
            except:
                pass


        if p.text.strip(' :\n').lower() == 'interests':
            try:
                for j in dc['Interests']:
                    doc.paragraphs[i+1].add_run('    •   ' + j.strip() + '\n')
    #                 doc.paragraphs[i+1].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            except:
                pass


    doc.save(path_save)
    print("Process has Completed...")

