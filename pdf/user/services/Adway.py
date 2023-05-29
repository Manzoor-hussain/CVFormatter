import os
import openai
import docx
import docx2txt
import re
import json
from .keys import api_key
from docx.enum.text import WD_UNDERLINE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
import PyPDF2
import pdfplumber


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]

def adway_converter(path ,formatted_path,save_path):
    formatted= formatted_path
#     un_formatted=os.getcwd() + path
    
#     doc = docx.Document(un_formatted)
#     formated_text = docx2txt.process(formatted)
#     unformated_text = docx2txt.process(un_formatted)
    def read_text_from_docx(path):
        doc = docx.Document(path)
        text = [paragraph.text for paragraph in doc.paragraphs]
        return '\n'.join(text)

    def read_text_from_pdf(path):
        with open(path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = []
            for page in pdf_reader.pages:
                text.append(page.extract_text())
            return '\n'.join(text)

    if path.endswith('.docx'):
        unformated_text = read_text_from_docx(path)
    elif path.endswith('.pdf'):
        unformated_text = read_text_from_pdf(path)
    else:
        unformated_text = 'Unsupported file format'
        

    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key
    # llm=OpenAI(temperature=0, max_tokens=1500,openai_api_key=api_key)
    fields_labels = "Name, PROFILE, EDUCATION, IT LITERACY, CERTIFICATES, PROJECTS, PROFESSIONAL QUALIFICATIONS, SOFTWARES, languages, Interests, TRAININGS, skills, WORK EXPERIENCE"

    
    print ("Process has Started...")
    test_text = """

    Ectract data from this text:

    \"""" + unformated_text + """\"

    in following JSON format:
{
"Name":"candidate name",
"Right to work":"value",
"Resides":"location of candidiate",
"Summary" : "value",

"Employment History" : [
    {"Duration" : "Working duration in company",
     "Company Name" : "Name of company",
     "Designation" : "Specific designation in that Company",
     "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
    },
    
    {"Duration" : "Working duration in company",
     "Company Name" : "Name of company",
     "Designation" : "Specific designation in that Company",
     "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
    },
    ...
    ]
"Education" : [
    {"Duration":"duration of that degree,
     "Institute Name":"Name of that institute",
     "Location":"Location of that institute "
     "Degree":"Name of that degree",
    },
    
    {"Duration":"duration of that degree,
     "Institute Name":"Name of that institute",
     "Location":"Location of that institute "
     "Degree":"Name of that degree",
    },
...],
"Skills" : ["skill1", "skill2", ...],
"Qualifications" : ["qualifications1", "qualifications2", ...],
"Certifications":["certifications1","certifications2",...]
"Languages":["languages1","languages2",...]
"Interests":["interests1","interests2",...]

}

    Do not include Grade

    Do not include Mobile number, Emali and home address 
    """


    result = get_completion(test_text)

    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):
        try:
            if p.text.strip().lower() == 'name':
    #                 doc.paragraphs[i].text = ""
                    run = doc.paragraphs[i].add_run(dc['Name'].strip().title())
                    run.bold = True
    #                 run.font.size = Pt(16.5)
        except:
            pass
    #     try:
    #         if p.text.strip().lower() == 'location':
    #                 doc.paragraphs[i].text = ""
    #                 run = doc.paragraphs[i].add_run(dc['Location'].strip().title())
    #                 run.bold = True
    # #                 run.font.size = Pt(16.5)
    #     except:
    #         pass
        try:
            if p.text.strip(' :\n').lower() == 'summary':
                doc.paragraphs[i+2].add_run(dc['Summary'].strip()).bold = False
                doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        except:
            pass
        try:
            if p.text.strip(' :\n').lower() == 'resides':
                run=doc.paragraphs[i].add_run("\t"+dc['Resides'].strip())
                run.bold=True

        except:
            pass

        if p.text.strip(' :\n').lower() == 'education':
            for j in dc['Education']:
                run1 = doc.paragraphs[i+2].add_run(j['Duration'].strip()+"\n")
                run1.bold=False
                run2 = doc.paragraphs[i+2].add_run("\t\t\t\t\t\t\t"+j['Institute Name'].strip()+"\n").bold=True
    #             run2 = doc.paragraphs[i+2].add_run("\t\t\t\t\t\t\t"+j['Location'].strip()+"\n").bold=True
                run4 = doc.paragraphs[i+2].add_run("\t\t\t\t\t\t\t"+j['Degree'].strip()+"\n")
                run4.bold=True
    #             run4.italic=True

    #             doc.paragraphs[i+2].add_run(j['Duration'].strip()+"\n").bold=False


    #             doc.paragraphs[i+2].add_run(j['Thesis'].strip() + "\n\n").bold=False

    #               + j['Institute'].strip() + "–" + j['Duration'].strip()).bold = False
    #               doc.paragraphs[i+2].add_run('  • ' + j.strip() + '\n').bold = False



        try:
            if p.text.strip(' :\n').lower() == 'certifications':
                for j in dc['Certifications']:
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
            if p.text.strip(' :\n').lower() == 'training':
                for j in dc['Training']:
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
            if p.text.strip(' :\n').lower() == 'employment history':
                for j in dc['Employment History']:
                    run1=doc.paragraphs[i+2].add_run(j['Duration'].strip() + "\n").bold=True
                    run2=doc.paragraphs[i+2].add_run("\t\t\t\t\t\t"+j['Company Name'].strip()+"\n").bold=True
                    run3=doc.paragraphs[i+2].add_run("\t\t\t\t\t\t"+j['Designation'].strip() + "\n\n")
                    run3.bold=True
    #                 run3.italic=True
    #                 run4=doc.paragraphs[i+2].add_run("\t\t\t\t\t\t"+j['Location'].strip()+"\n\n")
    #                 run4.bold=False
    #                 doc.paragraphs[i+2].add_run(j['Duration'].strip()+'\t\t'+j['Company Name'].strip()+ "\n").bold = True
    #                 doc.paragraphs[i+2].add_run('Responsibilities:' + '\n').bold = True
    #                 if j['Responsibilities']:
    #                     doc.paragraphs[i+2].add_run('\n')
                    for k in j['Responsibilities']:
                        doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n').bold = False
                    doc.paragraphs[i+2].add_run('\n\n')
        except:
            pass

    doc.save(save_path)
    print("\n")
    print("---------------------------------------------------------------------------------------------------------------------")
    print("\n")
    print("Process has Completed...")


