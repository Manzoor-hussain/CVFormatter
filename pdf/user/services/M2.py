import os
import re
import openai
import docx
import docx2txt
import traceback
import json
from docx.shared import Pt
import PyPDF2 
from .keys import api_key
from docx2python import docx2python


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]

def m2_partnership_converter(path,pathout,path_save):
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
            unformated_text = docx2txt.process(un_formatted)
            print('Its Docx')
        except:
            print('WE DONT SUPPORT THIS TYPE OF FILE')
    
    print("Proces has Started...")
    test_text = """
    Extract data from this text:

    \"""" + re.sub('\n+','\n', unformated_text) + """\"

    in following JSON format:
    {
    "Candidate Name" : "value",
    "Work Experience" : [
        {"Company Name" : "Name of company",
        "Company Location" : "Location of company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
        },
        {"Company Name" : "Name of company",
        "Company Location" : "Location of company",
        "Duration" : "Working Duration in Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...]
        },
        ...
        ],
    "Education" : [
        {"Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute"
        },
        {"Institute Name" : "Name Of institute",
        "Degree Name": "Name of degree",
        "Duration" : "Studying duration in institute"
        },
        ...
        ],
    "Awards" : ["award 1", "award 2", ...],
    "Skills" : ["skill 1", "skill 2", ...],
    "Projects" : ["project 1", "project 2", ...],
    "Languages" : ["language 1", "language 2", ...],
    "Interests" : ["interest 1", "interest 2", ...],
    "Trainings" : ["training 1", "training 2", ...],
    "Hobbies" : ["hobby 1", "hobby 2", ...]
    }
    """

    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key
    result = get_completion(test_text)

    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))

    # Open the existing document
    doc = docx.Document(formatted)

    for table in doc.tables:
#         print(table)
        for row in table.rows:
            for cell in row.cells:
                for tb in cell.tables:
                    for i,r in enumerate(tb.rows):
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'name':
                                tb.cell(i,0).text = ""
                                para = tb.cell(i,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                run = para.add_run(dc['Candidate Name'].strip() + '\n')
                                run.font.size = Pt(22)
                                run.bold = True
                        except:
                            print(traceback.print_exc())                  
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'e education':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Education']:
                                    try:
                                        if j['Institute Name'].strip():
                                            para.add_run(j['Institute Name'].strip() + '\n').bold = True
                                        else:
                                            para.add_run('Institute Name not mentioned\n').bold = True
                                    except:
                                        para.add_run('Institute Name not mentioned\n').bold = True                                 
                                    try:
                                        if j['Duration'].strip():
                                            para.add_run(j['Duration'].strip() + '\n').bold = True
                                        else:
                                            para.add_run('Duration not mentioned\n').bold = True
                                    except:
                                        para.add_run('Duration not mentioned\n').bold = True
                                    try:
                                        if j['Degree Name'].strip():
                                            para.add_run(j['Degree Name'].strip() + '\n\n').bold = False
                                        else:
                                            para.add_run('Degree Name not mentioned\n\n').bold = False
                                    except:
                                        para.add_run('Degree Name not mentioned\n\n').bold = False                           


                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 's skills':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Skills']:
                                    para.add_run('    • ' + j.strip() + '\n').bold = False


                        except:
                            print(traceback.print_exc())
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'h hobbies':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Hobbies']:
                                    para.add_run('    • ' + j.strip() + '\n').bold = False
                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'e experience':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Work Experience']:
                                    try:
                                        if j['Company Name'].strip():
                                            para.add_run(j['Company Name'].strip() + '\n').bold = True
                                        else:
                                            para.add_run('Company Name not mentioned\n').bold = True
                                    except:
                                        para.add_run('Company Name not mentioned\n').bold = True
                                    try:
                                        if j['Company Location'].strip():
                                            para.add_run(j['Company Location'].strip() + ' | ').bold = False
                                        else:
                                            para.add_run('Company Location not mentioned | ').bold = False
                                    except:
                                        para.add_run('Company Location not mentioned | ').bold = False
                                    try:
                                        if j['Duration'].strip():
                                            para.add_run(j['Duration'].strip() + '\n\n').bold = False
                                        else:
                                            para.add_run('Duration not mentioned\n\n').bold = False
                                    except:
                                        para.add_run('Duration not mentioned\n\n').bold = False
                                    try:
                                        if j['Responsibilities']:
                                            for k in j['Responsibilities']:
                                                para.add_run('    • ' + k.strip() + '\n').bold = False
                                    except:
                                        pass
                                    para.add_run('\n\n').bold = False                            


                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'k key projects':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Projects']:
                                    para.add_run(j.strip() + '\n').bold = False
                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'a awards':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Awards']:
                                    para.add_run(j.strip() + '\n').bold = False
                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'l languages':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Languages']:
                                    para.add_run('    • ' + j.strip() + '\n').bold = False
                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 'i interests':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Interests']:
                                    para.add_run('    • ' + j.strip() + '\n').bold = False
                        except:
                            pass 
                        try:
                            if r.cells[0].text.strip(' :\n').lower() == 't trainings':
                                para = tb.cell(i+1,0).add_paragraph()
                                para.style.font.name = "Century Gothic"
                                for j in dc['Trainings']:
                                    para.add_run(j.strip() + '\n').bold = False
                        except:
                            print(traceback.print_exc())                   



    # Save the updated document as a new file
    doc.save(path_save)
    print("Process has Completed...")
    

