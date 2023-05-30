import os
import openai
import docx
import docx2txt
import re
import json
import PyPDF2
from docx.shared import Pt
from docx.enum.text import WD_UNDERLINE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from .keys import api_key


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]


def edex_converter(path, pathout, path_save):
    formatted= pathout
    un_formatted = path
    formated_text = docx2txt.process(formatted)

    try:
        with open(un_formatted, 'rb') as file:
        # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(file)
            unformatted_text = ""
            for i in range (len(pdf_reader.pages)):
                first_page = pdf_reader.pages[i]
                unformatted_text += first_page.extract_text()
            print('Its PDF')
    except:
        try:
            unformatted_text = docx2txt.process(un_formatted)
            print('Its Docx')
        except:
            print('WE DONT SUPPORT THIS TYPE OF FILE')

    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key

    print("Process has Started...")

    test_text = """

    Ectract data from this text:

    \"""" + unformatted_text + """\"

    in following JSON format:
    {
    "Name" : "value",
    "Profile" : "value",

    "Education" : [
        {"Institute Name" : "Name Of institute",
        "Duration" : "Studying duration in institute",
        "Degree Name": "Name of degree",
        },
        {"Institute Name" : "Name Of institute",
        "Duration" : "Studying duration in institute",
        "Degree Name": "Name of degree",
        },
        ...
    ],
    "It Literacy" : ["literacy1", "literacy2", ...],
    "Certificates" : ["certificate1", "certificate2", ...],
    "Projects" : ["project1", "project2", ...],
    "Professional Qualifications" : ["qualification1", "qualification2", ...],
    "Softwares" : ["software1", "software2", ...],
    "Languages" : ["language1", "language2", ...],
    "Interests" : ["interest1", "interest2", ...],
    "Trainings" : ["training1", "training2", ...],
    "Skills" : ["skill1", "skill2", ...],
    "Work Experience" : [
        {"Company Name" : "Name of company",
        "Duration" :  "Working Duration in Company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        {"Company Name" : "Name of company",
        "Duration" :  "Working Duration in Company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        ...
        ]
    }

    Please keep the following points in considration while extracting data from text:
        1. Do not summarize or rephrase Responsibilities. Extract each Responsibility completely from text.
        2. Make it sure to keep the response in JSON format.
        3. If value not found then leave it empty/blank.
        4. Do not include Mobile number, Email and Home address.
        5. Do not include Grade  
    """


    result = get_completion(test_text)
    
    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):

        try:
            if p.text.strip().lower() == 'name:':
                doc.paragraphs[i].text = ""
                run = doc.paragraphs[i].add_run(dc['Name'].strip().title())
                run.bold = True
                run.font.size = Pt(14)
        except:
               pass
        try:
            if p.text.strip(' :\n').lower() == 'profile':
                doc.paragraphs[i+2].add_run(dc['Profile'].strip()).bold = False
                doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        except:
            pass

        try:        
            if p.text.strip(' :\n').lower() == 'education':
                for j in dc['Education']:
        #             doc.paragraphs[i+2].add_run(j['Institute Name']).bold = Fals
                    doc.paragraphs[i+2].add_run(j["Institute Name"].strip() + ' – ' + j["Duration"].strip() + '\n').font.underline = True
                    doc.paragraphs[i+2].add_run(j['Degree Name'].strip()).bold = True
                    doc.paragraphs[i+2].add_run("\n\n")
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'it literacy':
                for j in dc['It Literacy']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'certificates':
                for j in dc['Certificates']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'projects':
                for j in dc['Projects']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'qualifications':
                for j in dc['Qualifications']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'professional qualifications':
                for j in dc['Professional Qualifications']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'softwares':
                for j in dc['Softwares']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'languages':
                for j in dc['Languages']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'interests':
                for j in dc['Interests']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'trainings':
                for j in dc['Trainings']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'skills':
                for j in dc['Skills']:
                    doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'work experience':
                for j in dc['Work Experience']:
                    doc.paragraphs[i+2].add_run(j['Company Name'].strip() + ' – ' + j['Duration'] + '\n').font.underline = True
                    doc.paragraphs[i+2].add_run(j['Designation'].strip()).bold = True
                    doc.paragraphs[i+2].add_run('\n\n')
                    if len(j["Responsibilities"]) == 0:
                        pass
                    else:
                        len(j["Responsibilities"]) != 0
                        doc.paragraphs[i+2].add_run("Responsibilities:" + '\n').bold = True
                        for k in j['Responsibilities']:
                            doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n').bold = False
                    doc.paragraphs[i+2].add_run('\n')
        except:
            pass

    doc.save(path_save)
    print("Process has Completed...")