import os
import openai
import docx
import docx2txt
import re
import json
from .keys import api_key
from docx.enum.text import WD_UNDERLINE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def get_completion(prompt, model="gpt-3.5-turbo"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0, # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]

def alexander_steele_converter(path,pathoutput,save_path):
    formatted = pathoutput
    # un_formatted=os.getcwd() + "/Unformated_EdEx/Adil Thomas Daraki CV .docx"
    # un_formatted=os.getcwd() + "/Unformated_EdEx/Alice Maynard CV.docx"
    #un_formatted=os.getcwd() + path
    # un_formatted=os.getcwd() + "/Unformated_EdEx/Amrit Bassan CV.docx"
    
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

# Example usage
    if path.endswith('.docx'):
        unformated_text = read_text_from_docx(path)
    elif path.endswith('.pdf'):
        unformated_text = read_text_from_pdf(path)
    else:
        print('Unsupported file format')
    
    os.environ["OPEN_API_KEY"] = api_key
    openai.api_key = api_key
    # llm=OpenAI(temperature=0, max_tokens=1500,openai_api_key=api_key)
    fields_labels = "Name, PROFILE, EDUCATION, IT LITERACY, CERTIFICATES, PROJECTS, PROFESSIONAL QUALIFICATIONS, SOFTWARES, languages, Interests, TRAININGS, skills, WORK EXPERIENCE"

    
    print ("Process has Started...")
    test_text = """
    Extract data from this text:

    \"""" + unformated_text + """\"

    in following JSON format:
    {
    "Profile Summary" : "value",
    "Experience/Employment History" : [
        {"Company Name" : "Name of company",
        "Duration" : "Working Duration in Company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        {"Company Name" : "Name of company",
        "Duration" : "Working Duration in Company",
        "Designation" : "Specific designation in that Company",
        "Responsibilities" : ["Responsibility 1", "Responsibility 2", ...],
        },
        ...
        ],
    "Education" : [
        {"Institute Name" : "Name Of institute",
        "Duration" : "Studying duration in institute",
        "Degree title" : "Name or title of Degree",
        },
        {"Institute Name" : "Name Of institute",
        "Duration" : "Studying duration in institute",
        "Degree title" : "Name or title of Degree",
        },
        ...
        ],
    "Professional Training/Courses" : ["Training1", "Training2", ...],
    "Achievements" : ["achievement1", "achievement2", ...],
    "Languages" : ["language1", "language2", ...],
    "Interests" : ["interest1", "interest2", ...],
    "Computer Skills" : ["computerskill1", "computerskill2", ...],
    "Skills" : ["skill1", "skill2", ...],

    }

    Do not include Grade

    Do not return those keys against which no value will be founded

    Do not include Mobile number, Emali and home address
    """

    result = get_completion(test_text)

    print(result)
    dc = dict(json.loads(re.sub(',[ \n]*\]',']',re.sub(',[ \n]*\}','}',result.replace('...','')))))

    doc = docx.Document(formatted)

    for i,p in enumerate(doc.paragraphs):
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
                    doc.paragraphs[i+2].add_run('  • ' + j["Duration"].strip() + ' – ' + j["Institute Name"].strip() + ', ' + j["Degree title"].strip() + '\n')
        except:
            pass

        try:
            if p.text.strip(' :\n').lower() == 'professional training':
                for j in dc['Professional Training/Courses']:
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
            if p.text.strip(' :\n').lower() == 'computer skills':
                for j in dc['Computer Skills']:
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


        if p.text.strip(' :\n').lower() == 'experience':
            for j in dc['Experience/Employment History']:

                a = ceil(len(j["Duration"]) * 1.75)
                print(a)
                b = " " * a
    #             c = floor(a / 4)
    #             if a % 4 == 3:
    #                 b = " " * 1 
    #             elif a % 4 == 2:
    #                 b = " " * 2
    #             else:
    #                 a % 4 == 1
    #                 b = " " * 3

                run = doc.paragraphs[i + 2].add_run(j["Duration"].strip()+ '\t\t')
                run.font.bold = False  # Unbold the first company name

                run = doc.paragraphs[i + 2].add_run(j['Company Name'] + '\n')
                run.font.bold = True
    #             doc.paragraphs[i+2].add_run(j["Duration"].strip() +"\t\t" + j["Company Name"].strip() + "\n")
                doc.paragraphs[i+2].add_run(b + "\t\t" + j["Designation"].strip() + "\n").bold=True
                for k in j["Responsibilities"]:
                    doc.paragraphs[i+2].add_run('  • ' + k.strip() + "\n")
                doc.paragraphs[i+2].add_run('\n\n')

    #     except:
    #         pass

    doc.save(save_path)
    print("\n")
    print("---------------------------------------------------------------------------------------------------------------------")
    print("\n")
    print("Process has Completed...")
