from openpyxl.reader.excel import load_workbook
from pywebio.input import *
from pywebio.output import *
from pywebio.pin import *
from pywebio import start_server
import os
from fpdf import FPDF
import pandas as pd
from openpyxl import Workbook
os.system('cls')

def app():
    pass
    # input form
    info = input_group("Upload CSV files", [
        file_upload("Upload Roll Number Sheet", name="file1", required=True, accept=".csv"),
        file_upload("Upload Subject Master Sheet",  name="file2", required=True, accept=".csv"),
        file_upload("Upload Gardes Sheet",  name="file3", required=True, accept=".csv"),
        file_upload("Upload SEAL", name="img1", accept="image/*", multiple=False, required=False),
        file_upload("Upload Signature", name="img2", accept="image/*", multiple=False, required=False)
    ])

    file1 = info['file1']
    file2 = info['file2']
    file3 = info['file3']
    img1 = info['img1']
    img2 = info['img2']

    # saving names-roll.csv
    content1 = file1['content'].decode('utf-8').splitlines()
    save_csv(content1, "names-roll")

    # saving subjects_master.csv 
    content2 = file2['content'].decode('utf-8').splitlines()
    save_csv(content2, "subjects_master")

    # saving grades.csv 
    content3 = file3['content'].decode('utf-8').splitlines()
    save_csv(content3, "grades")

    # saving both seal and signature
    if img1:
        img1["filename"] = "SEAL."+img1["filename"].split('.')[1]
        open(r'input/'+img1["filename"], 'wb').write(img1['content'])
    # elif os.path.exists(r"input/SEAL")==True:
    #     os.remove(r"input/SEAL")
    
    if img2:
        img2["filename"] = "Signature."+img2["filename"].split('.')[1]
        open(r'input/'+img2["filename"], 'wb').write(img2['content']) 
    # elif os.path.exists(r"input/Signature")==True:
    #     os.remove(r"input/Signature")

    # Generating all xlsx files
    generate_marksheets()

    # Format required by the user
    put_input("range", label="Type the range of Roll Numbers")
    put_buttons(["Generate Range Transcripts","Generate All Transcripts"],onclick=[generate_range_transcripts, generate_all_transcripts])


def generate_marksheets():
    roll_num = pd.read_csv(r"input//names-roll.csv")
    subjects = pd.read_csv(r"input//subjects_master.csv", index_col=0)
    grades = pd.read_csv(r"input//grades.csv")

    name_dict={}
    # for index, row in roll_num.iterrows():
    #     name_dict[row['Roll']]=row['Name']
    print(roll_num)
    for i in range(len(roll_num)):
        key=roll_num.loc[i,'Roll']
        val=roll_num.loc[i,'Name']
        name_dict[key]=val

    if os.path.exists(r"output//")==False:
        os.mkdir(r"output//")

    heading = ("Sl No.", "Subject No.",	"Subject Name",	"L-T-P", "Credit", "Subject Type", "Grade")
    rolls = {}

    for i,row in grades.iterrows():

        roll = row["Roll"].upper()
        subcode = row["SubCode"]
        sem = row["Sem"]
        credit = row["Credit"]
        grade = row["Grade"]
        subtype = row["Sub_Type"]

        if roll not in rolls:
            rolls[roll] = {}
            rolls[roll]["Sem"] = []

        pdf = FPDF('L', 'mm', 'A3')
        pdf.add_page()
        pdf.set_font('Arial',size= 12)
        pdf.image('iitp_logo.jpeg', x = None, y = None, w = 0, h = 0, type = 'jpeg')
        pdf.multi_cell(200,10,"Roll No:  %s      Name:  %s      Year of Admission:  %s" %(roll,name_dict[roll],'20'+roll[:2]),border=1,align='C')
        
        path = 'TranscriptIITP/' + roll + '.pdf'
        pdf.output(path, 'F')

        # if os.path.exists(r"output//"+'.xlsx')==False:
        #     wb = Workbook()
        #     sheet = wb.active
        #     sheet.title = "Overall"
        #     sheet.append(["Roll No.", roll])
        #     sheet.append(["Name of Student", roll_num["Name"].loc[roll]])
        #     sheet.append(["Discipline", roll[4]+roll[5]]) 
        # else:
        #     wb = load_workbook(r"output//"+roll+'.xlsx')

    #     if sem not in rolls[roll]["Sem"]:
    #         rolls[roll]["Sem"].append(sem)
    #         rolls[roll]["Credit"+str(sem)] = 0
    #         rolls[roll]["Marks"+str(sem)] = 0

    #     rolls[roll]["Credit"+str(sem)] += int(credit)
    #     rolls[roll]["Marks"+str(sem)] += grade_to_marks(grade) * int(credit)
        
    #     if "Sem"+str(sem) not in wb.sheetnames:
    #         wb.create_sheet(title="Sem"+str(sem))
    #         sheet = wb["Sem"+str(sem)]
    #         sheet.append(heading)
    #         ind = 1
    #     else:
    #         sheet = wb["Sem"+str(sem)]
    #         ind = int(sheet.cell(row=sheet.max_row, column=1).value)+1

    #     sheet.append([ind, subcode, subjects["subname"].loc[subcode], subjects["ltp"].loc[subcode], credit, subtype, grade])
    #     wb.save(r"output//"+roll+'.xlsx')

    # for roll in rolls:
    #     wb = load_workbook(r"output//"+roll+".xlsx")
    #     sheet = wb["Overall"]

    #     rolls[roll]["Sem"].sort()
    #     sheet.append(["Semester No."]+rolls[roll]["Sem"])

    #     row1 = ["Semester wise Credit Taken"]
    #     spi = ["SPI"]
    #     row2 = ["Total Credits Taken"]
    #     cpi = ["CPI"]
    #     ind = 0
    #     for i in rolls[roll]["Sem"]:
    #         row1.append(rolls[roll]["Credit"+str(i)])
    #         spi.append(round(rolls[roll]["Marks"+str(i)]/rolls[roll]["Credit"+str(i)],2))

    #         if ind != 0:
    #             rolls[roll]["Marks"+str(i)] += rolls[roll]["Marks"+str(ind)]
    #             rolls[roll]["Credit"+str(i)] += rolls[roll]["Credit"+str(ind)]

    #         row2.append(rolls[roll]["Credit"+str(i)])
    #         cpi.append(round(rolls[roll]["Marks"+str(i)]/rolls[roll]["Credit"+str(i)],2))

    #         ind += 1

    #     sheet.append(row1)
    #     sheet.append(spi)
    #     sheet.append(row2)
    #     sheet.append(cpi)
    #     wb.save(r"output//"+roll+'.xlsx')



def grade_to_marks(grade):
    if 'AA' in grade:
        return 10
    elif 'AB' in grade:
        return 9
    elif 'BB' in grade:
        return 8
    elif 'BC' in grade:
        return 7
    elif 'CC' in grade:
        return 6
    elif 'CD' in grade:
        return 5
    elif 'DD' in grade:
        return 4
    elif 'F' in grade or 'I' in grade:
        return 0
    
    print(grade)


def generate_range_transcripts():
    pass


def generate_all_transcripts():
    
    pass


# function to save the content of the uploaded csv by the filename provided
def save_csv(content: list, filename):
    if os.path.exists(r"input//")==False:
        os.mkdir(r"input//")

    with open(r"input//" + '%s.csv' %filename, "w+") as csv_file:
        for line in content:
            csv_file.write(line + "\n")


if __name__=='__main__':
    generate_marksheets()
    start_server(app, port=36536, debug=True)
