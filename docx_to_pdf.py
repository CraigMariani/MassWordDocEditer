from docx2pdf import convert
from pathlib import Path
import os
import subprocess

if __name__ == '__main__':
    files = os.listdir()
    new_pdfs = []
    unchanged = []

    # create resultr directory 
    if os.path.isdir('pdfs') == False:
       subprocess.call(['mkdir', 'pdfs'])

    # iterate through files in current working directory
    for file in files:
        file_ext = Path(file).suffix
        
        # filter all docx files
        if file_ext == '.docx':
            file_name = os.path.splitext(file)[0]

            # file_path=os.getcwd()

            # convert the file to a pdf and store it in pdf directory
            convert(file, 'pdfs/{}.pdf'.format(file_name))
            new_pdfs.append(file)
        else:
            unchanged.append(file)
    
    print('New_PDFS: {}'.format(new_pdfs))
    print('Unchanged: {}'.fomrat(unchanged))