import re
import os
import magic
import string
import configparser
from bs4 import BeautifulSoup
import openpyxl

def readtext(path):
    path = path.rstrip()
    path = path.replace(' \n', '')
    path = path.replace('%0A', '')
    
    relative_path = path
    filename = os.path.abspath(relative_path)
    
    filetype = magic.from_file(filename, mime=True)
    print(f"\nEmbedding {filename} as {filetype}")
    text = ""

    if filetype == 'application/pdf':
        print('PDF not supported yet')
    elif filetype == 'text/plain':
        with open(filename, 'rb') as f:
            text = f.read().decode('utf-8')
    elif filetype == 'text/html':
        with open(filename, 'rb') as f:
            soup = BeautifulSoup(f, 'html.parser')
            text = soup.get_text()
    elif filetype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        text = read_excel(filename)

    return text

def read_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    text = ""
    for row in sheet.iter_rows(values_only=True):
        text += ' '.join(map(str, row)) + '\n'
    return text

def getconfig():
    config = configparser.ConfigParser()
    config.read('config.ini')
    return {section: dict(config.items(section)) for section in config.sections()}

if __name__ == "__main__":
    file_path = "categorized_data.xlsx"
    text_content = readtext(file_path)
    print(text_content)