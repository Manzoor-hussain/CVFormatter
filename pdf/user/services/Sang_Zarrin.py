import os
import openai
import docx
import re
import json
from pprint import pprint
import docx2txt
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

def sang_zarrin_converter(path, pathout, path_save):
    formatted = pathout
    un_formatted = path
    formated_text = docx2txt.process(formatted)

    try:
        with open(un_formatted, 'rb') as file:
        # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(file)
            unformated_text = ""
            for i in range (len(pdf_reader.pages)):
                first_page = pdf_reader.pages[i]
                unformated_text += first_page.extract_text()
            print('Its PDF')
    except:
        try:
    #         un_formatted.split(".")[-1] == "docx"
            unformated_text = docx2txt.process(un_formatted)
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
    "Summary" : "value",
    "Experience" : [
        {"Designation" : "The specific designation or position on which he works in this company",
        "Company Name" : "Name of company",
        "Location" : "Location of company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
        },
        {"Designation" : "The specific designation or position on which he works in this company",
        "Company Name" : "Name of company",
        "Location" : "Location of company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
        },
        ...
        ],
    "Education" : [
        {"Institute Name" : "Name Of institute and its location if available separated with comma ",
        "Institute locality" : "Location of instituion",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute"
        },
        {"Institute Name" : "Name Of institute and its location if available separated with comma ",
        "Institute locality" : "Location of instituion",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute"
        },
        ...
        ],
    "Trainings" : ["trainings1", "trainings2", ...],
    "Computer Skills" : ["computer skill1", "computer skill2", ...],
    "Skills" : ["skill1", "skill2", ...],
    "Qualifications" : ["qualification1", "qualification2", ...],
    "Languages" : ["language1", "language2", ...],
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

    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):
        try:
            if p.text.strip(' :\n').lower() == 'summary':
                doc.paragraphs[i+2].add_run(dc['Summary'].strip()).bold = False
                doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'experience':
                for j in dc['Experience']:
                    doc.paragraphs[i+2].add_run(j['Designation'].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j["Company Name"].strip() + ',  ' + j["Location"].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n\n').bold = True
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

        try:
            if p.text.strip(' :\n').lower() == 'education':
                for j in dc['Education']:
                    doc.paragraphs[i+2].add_run(j['Institute Name'].strip() +", "+j['Institute locality'] + '\n').bold = False
                    doc.paragraphs[i+2].add_run(j['Degree Nmae'].strip() + '\n').bold = False
                    doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n\n').bold = False
        except:
            pass

        try: 
            if p.text.strip(' :\n').lower() == 'trainings':
                for j in dc['Trainings']:
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
            if p.text.strip(' :\n').lower() == 'skills':
                for j in dc['Skills']:
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

    doc.save(path_save)
    print("Process has Completed...")
    
    