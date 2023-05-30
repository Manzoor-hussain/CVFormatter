import os
import openai
import docx
from pprint import pprint
import docx2txt
import PyPDF2
import re
import json
from .keys import api_key

def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]


def aspion_converter(path,pathout,path_save):
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


    openai.api_key = api_key
    
    print("Process has Started...")
    test_text = """Extract data from this text:

    \"""" + unformated_text + """\"

    in following JSON format:
    {
    "Profile Summary" : "value",

    "Education" : [
        {"Institute" : "Name Of institute",
        "Duration": "Studying duration in institute",
        "Location" : "Location of institute",
        "Degree title" : "Name of degree completed from that institute"
        },
        {"Institute" : "Name Of institute",
        "Duration": "Studying duration in institute",
        "Location" : "Location of institute",
        "Degree title" : "Name of degree completed from that institute"
        },
        ...
        ],
    "Achievements" : ["achievement1", "achievement2", ...],
    "Languages" : ["language1", "language2", ...],
    "Interests" : ["interest1", "interest2", ...],
    "Trainings" : ["training1", "training2", ...],
    "Skills" : ["skill1", "skill2", ...],
    "Experience" : [
        {"Designation" : "Specific designation in that Company",
        "Name of Company" : "Company Name here",
        "Company locality" : "Location of company",
        "Duration" : "Time period in whic he the person has worked in that company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        {"Designation" : "Specific designation in that Company",
        "Name of Company" : "Company Name here",
        "Company locality" : "Location of company",
        "Duration" : "Time period in whic he the person has worked in that company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        ...
        ]
    }

    You must keep the following points in considration while extracting data from text:
        1. Do NOT split, rephrase or summarize list of Responsibilities. Extract each Responsibility as a complete sentence from text.
        2. Make it sure to keep the response in JSON format.
        3. If value not found then leave it empty/blank.
        4. Do not include Mobile number, Email and Home address.
    """


    result = get_completion(test_text)

    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):    

        try:
            if p.text.strip(' :\n').lower() == 'summary':
                doc.paragraphs[i+2].add_run(dc['Summary'].strip()).bold = False
                doc.paragraphs[i+2].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        except:
            pass

        try:     
            if p.text.strip(' :\n').lower() == 'education':
                for j in dc['Education']:
        #             doc.paragraphs[i+2].add_run(j['Institute Name']).bold = Fals
                    doc.paragraphs[i+2].add_run(j['Institute'].strip() + ", " + j["Location"].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j['Degree title'].strip() + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n\n').bold = True

        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'achievements':
                for j in dc['Achievements']:
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
            if p.text.strip(' :\n').lower() == 'career history':
                for j in dc['Experience']:
                    doc.paragraphs[i+2].add_run(j['Designation'].strip() + '\n').bold = True
                    if j ['Company locality'] != "":
                        doc.paragraphs[i+2].add_run(j['Company locality'].strip() + '\n').bold = True
                    else:
                        j ['Company locality'] == ""
                        doc.paragraphs[i+2].add_run("N/A" + '\n').bold = True
                    doc.paragraphs[i+2].add_run(j['Duration'].strip() + '\n\n').bold = True
                    if len(j["Responsibilities"]) == 0:
                        pass
                    else:
                        j["Responsibilities"] != 0
                        doc.paragraphs[i+2].add_run("Responsibilities:" + '\n').bold = True
                        for k in j['Responsibilities']:
                            doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n').bold = False
                        doc.paragraphs[i+2].add_run('\n')
        except:
            pass

    doc.save(path_save)
    print("Process has Completed...")
