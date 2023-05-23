import os
import openai
import docx
import docx2txt
import re
import json
import PyPDF2
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


def hbd_converter(path, pathout, path_save):
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

    \"""" + unformated_text + """\"

    in following JSON format:
    {
    "Name" : "value",
    "Profile" : "value",

    "Education" : [
        {"Institute" : ["Studying duration in institute", "Name Of institute", "Name of degree"]},
        {"Institute" : ["Studying duration in institute", "Name Of institute", "Name of degree"]},
        ...
        ],
    "Certificates" : ["certificate1", "certificate2", ...],
    "Achievements" : ["achievement1", "achievement2", ...],
    "Qualifications" : ["qualification1", "qualification2", ...],
    "Computer Skills" : ["computer skill1", "computer skill2", ...],
    "Expertise" : ["expertise1", "expertise2", ...],
    "Languages" : ["language1", "language2", ...],
    "Interests" : ["interest1", "interest2", ...],
    "Trainings" : ["training1", "training2", ...],
    "Skills" : ["skill1", "skill2", ...],
    "Work Experience" : [
        {"Name of Company" : ["Specific designation in that Company", "Name of Company"],
        "Duration" : "Working duration in that compnay",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        {"Name of Company" : ["Specific designation in that Company", "Name of Company"],
        "Duration" : "Working duration in that compnay",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        ...
        ]
    }

    Do not include Grade

    Do not include Mobile number, Emali and home address 
    """


    result = get_completion(test_text)
    
    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))
    
    doc = docx.Document(formatted)
    for i,p in enumerate(doc.paragraphs):
    
        try:        
            if p.text.strip(' :\n').lower() == 'name':
                doc.paragraphs[i].text = ""
                doc.paragraphs[i].add_run(dc["Name"].strip()).bold = True; doc.paragraphs[i].runs[-1].font.size = Pt(30)

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
                    doc.paragraphs[i+2].add_run('  • ' + j["Institute"][0].strip() + ' – ' + j["Institute"][1].strip() + ' – ' + j["Institute"][2].strip() + '\n' ).bold = False

        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'expertise':
                for j in dc['Expertise']:
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
            if p.text.strip(' :\n').lower() == 'achievements':
                for j in dc['Achievements']:
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
            if p.text.strip(' :\n').lower() == 'computer skills':
                for j in dc['Computer Skills']:
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
                    doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j['Name of Company'][0].strip() + ' – ' + j['Name of Company'][1] + '\n\n').bold = True
                    for k in j['Responsibilities']:
                        doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n').bold = False
                    doc.paragraphs[i+2].add_run('\n')
        except:
            pass

    doc.save(path_save)
    print("Process has Completed...")