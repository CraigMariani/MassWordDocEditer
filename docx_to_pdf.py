'''
How to run this file:

-install docx2pdf with "pip install docx2pdf"

-once it is installed run the script you will get an error like 
    " \Python\Python38\lib\site-packages\docx2pdf\__init__.py , line 106 in convert"

-go to that line where it says "return windows(paths, keep_active)" delete it and type:
        try:
            return windows(paths, keep_active)
        except AttributeError:
            print('attribute error init.py')

-There will be man errors in the script but it will work. The errors are probably from the module docx2pdf

'''


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
            try:
                convert(file, 'pdfs/{}.pdf'.format(file_name))
                new_pdfs.append(file)
            except AttributeError:
                print 'file {} unchanged'.format(file)
                unchanged.append(file)

            
        else:
            unchanged.append(file)
    
    print('New_PDFS: {}'.format(new_pdfs))
    print('Unchanged: {}'.fomrat(unchanged))