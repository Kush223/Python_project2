from pywebio.input import *
from pywebio.output import *
from pywebio.pin import *
from pywebio import start_server
import os
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

    # Format required by the user
    put_input("range", label="Type the range of Roll Numbers")
    put_buttons(["Generate Range Transcripts","Generate All Transcripts"],onclick=[generate_range_transcripts, generate_all_transcripts])


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
    start_server(app, port=36535, debug=True)
