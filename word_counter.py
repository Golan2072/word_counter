# Word counter script
# By Omer Golan-Joel
# v1.0 March 13th, 2020
# Open source software
# You can contact me at golan2072@gmail.com

import os
import openpyxl
import platform
import docx2txt
import PyPDF2


def current_dir():
    if platform.system() == "Windows":
        directory = os.listdir(".\\")
    else:
        directory = os.getcwd()
    return directory


def excel_counter(filename):
    count = 0
    wb = openpyxl.load_workbook(filename)
    for sheet in wb:
        for row in sheet:
            for cell in row:
                text = str(cell.value)
                if text != "None":
                    word_list = text.split()
                    count += len(word_list)
    return count


def pdf_counter(filename):
    pdf_word_count = 0
    pdfFileObj = open(filename, "rb")
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    number_of_pages = pdfReader.getNumPages() - 1
    for page in range(0, number_of_pages + 1):
        page_contents = pdfReader.getPage(page - 1)
        raw_text = page_contents.extractText()
        text = raw_text.encode('utf-8')
        page_word_count = len(text.split())
        pdf_word_count += page_word_count
    return pdf_word_count


def main():
    word_count = 0
    print(f"Current Directory: {os.getcwd()}")
    for file in current_dir():
        file_name_list = os.path.splitext(file)
        extension = file_name_list[1]
        if extension == ".xlsx":
            current_count = excel_counter(file)
            print(f"{file} {current_count}")
            word_count += current_count
        if extension == ".docx":
            text = docx2txt.process(file)
            current_count = len(text.split())
            print(f"{file} {current_count}")
            word_count += current_count
        if extension == ".txt":
            f = open(file, "r")
            text = f.read()
            current_count = len(text.split())
            print(f"{file} {current_count}")
            word_count += current_count
        if extension == ".pdf":
            pdf_word_count = pdf_counter(file)
            print(f"{file} {pdf_word_count}")
            word_count += pdf_word_count
        else:
            pass
    print(f"Total: {word_count}")


main()
