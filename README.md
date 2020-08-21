# MassWordDocEditer
Task automation script that can automate changes over hundreds of microsoft word documents


REQUIRED SOFTWARE:

 - Python3
 - git bash command line 
 - docx python module (type following command into command line to install: pip install python-docx)
 - subprocess python module (type following command into command line to install: pip install subprocess)


INSTRUCTIONS FOR RUNNING SCRIPT:

 - This script will only work with files of the .docx extension so be sure to save files like this 
 - Script will also need to be in the same directory as the files you are trying to change
 - The script iterates through all files in the directory that it is in and saves it in a new folder called "results_documents"
 - This folder is automitcally created if it does not exist
 - The script will prompt user for what change they want to make 
    Example: 
      Say if I wanted to change the date from 2019 to 2020 on all pages and footers type the following after the prompt
      
      2019 2020
