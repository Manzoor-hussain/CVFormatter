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

def ashbys_converter(path, pathout, path_save):
    
    formatted= pathout
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
    "Profile" : "value",

    "Education" : [
        {"Duration" : "Studying duration in institute",
        "Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        },
        {"Duration" : "Studying duration in institute",
        "Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        },
        ...
        ],
    "Work History" : [
        {"Duration" : "Working duration in specific company",
        "Company Name" : "Name of company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...],
        },
        {"Duration" : "Working duration in specific company",
        "Company Name" : "Name of company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...],
        },
        ...
        ],

    "Trainings" : ["training1", "training2", ...],
    "Projects" : [
        {"Project Name" : "tile/name of project",
        "Company Name" : "Name of company in which this project has done",
        "Designation" : "Specific designation for that project",
        "Responsibilities" : ["Responsibility1", "Responsibility2", ...],
        },
        {"Project Name" : "tile/name of project",
        "Company Name" : "Name of company in which this project has done",
        "Designation" : "Specific designation for that project",
        "Accountibilities" : ["accountibility1", "accountibility2", ...],
        },
        ...
        ],
    "Skills" : ["skill1", "skill2", ...],
    "Languages" : ["language1", "language2", ...],
    "Personal Skills" : ["personal skill1", "personal skill2", ...],
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
    print(result)
    dc = dict(json.loads(re.sub(r'\[\"\"\]',r'[]',re.sub(r'\"[Un]nknown\"|\"[Nn]one\"|\"[Nn]ull\"',r'""',re.sub(r',[ \n]*\]',r']',re.sub(r',[ \n]*\}',r'}',result.replace('...','')))))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):
        try:
            if p.text.strip(' :\n').lower() == 'profile':
                doc.paragraphs[i].text = ""
                doc.paragraphs[i].add_run(dc['Profile'].strip()).bold = False
                doc.paragraphs[i].alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        except:
            pass

        try:        
            if p.text.strip(' :\n').lower() == 'education':
                doc.paragraphs[i].text = ""
                for j in dc['Education']:
        #             doc.paragraphs[i+2].add_run(j['Institute Name']).bold = Fals
                    doc.paragraphs[i].add_run(j["Duration"].strip() + '\n').bold = False
                    doc.paragraphs[i].add_run(j["Institute Name"].strip() + '\n').bold = False
                    doc.paragraphs[i].add_run(j['Degree Name'].strip() + '\n\n').bold = True
        except:
            pass


        try:
            if p.text.strip(' :\n').lower() == 'projects':
                doc.paragraphs[i].text = ""
                for j in dc['Projects']:
                    doc.paragraphs[i].add_run(j['Project Name'].strip() + '\n').bold = True
                    doc.paragraphs[i].add_run(j['Company Name'].strip() + '\n').bold = True
                    doc.paragraphs[i].add_run(j['Designation'].strip() + '\n\n').bold = True
                    if len(j["Responsibilities"]) == 0:
                        pass
                    else:
                        len(j["Responsibilities"]) != 0
                        doc.paragraphs[i].add_run('Key Accountibilities:' + '\n').bold = False
                        for k in j['Responsibilities']:
                            doc.paragraphs[i+2].add_run('  • ' + k.strip() + '\n').bold = False
                        doc.paragraphs[i+2].add_run('\n')
        except:
            pass


        try:
            if p.text.strip(' :\n').lower() == 'personal skills':
                doc.paragraphs[i].text = ""
                for j in dc['Personal Skills']:
                    doc.paragraphs[i].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass


        try:
            if p.text.strip(' :\n').lower() == 'languages':
                doc.paragraphs[i].text = ""
                for j in dc['Languages']:
                    doc.paragraphs[i].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'interests':
                doc.paragraphs[i].text = ""
                for j in dc['Interests']:
                    doc.paragraphs[i].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'training':
                doc.paragraphs[i].text = ""
                for j in dc['Trainings']:
                    doc.paragraphs[i].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'skills':
                doc.paragraphs[i].text = ""
                for j in dc['Skills']:
                    doc.paragraphs[i].add_run('  • ' + j.strip() + '\n').bold = False
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'work history':
                doc.paragraphs[i].text = ""
                for j in dc['Work History']:
                    doc.paragraphs[i].add_run(j['Duration'].strip() + '\n').bold = True
                    doc.paragraphs[i].add_run(j['Company Name'].strip() + '\n').bold = True
                    doc.paragraphs[i].add_run(j['Designation'].strip() + '\n\n').bold = True
                    for k in j['Responsibilities']:
                        doc.paragraphs[i].add_run('\t' + '  • ' + k.strip() + '\n').bold = False
                    doc.paragraphs[i].add_run('\n')
        except:
            pass

    doc.save(path_save)
    print("Process has Completed...")
   