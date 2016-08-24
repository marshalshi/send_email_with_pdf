# -*- coding: utf-8 -*-
import os
import re
import shutil
from StringIO import StringIO

from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.pdf import ContentStream, TextStringObject
from PyPDF2.utils import isString, b_, u_, ord_, chr_, str_, formatWarning
from openpyxl import load_workbook

from django.conf import settings
from django.core.mail import EmailMessage

def convert_page_to_text(page):
    text = u_("")
    content = page.getContents()
    if not isinstance(content, ContentStream):
        content = ContentStream(content, page.pdf)
    # Note: we check all strings are TextStringObjects.  ByteStringObjects
    # are strings where the byte->string encoding was unknown, so adding
    # them to the text here would be gibberish.
    for operands, operator in content.operations:
        if operator == b_("Tj"):
            _text = operands[0]
            if isinstance(_text, TextStringObject):
                text += _text + ' '
        elif operator == b_("T*"):
            text += "\n"
        elif operator == b_("'"):
            text += "\n"
            _text = operands[0]
            if isinstance(_text, TextStringObject):
                text += operands[0] + ' '
        elif operator == b_('"'):
            _text = operands[2]
            if isinstance(_text, TextStringObject):
                text += "\n"
                text += _text + ' '
        elif operator == b_("TJ"):
            for i in operands[0]:
                if isinstance(i, TextStringObject):
                    text += i + ' '
            text += "\n"
            
    return text

def split_pdf(f_path, split_key_word):
    '''
    this function will split pdf to several sub-pdfs according key words
    which mark down to be split key.

    Args:
      f_path: target pdf path.
    '''

    pdf_folder = settings.PDF_FOLDER

    # get file name without extension
    f_base = os.path.basename(f_path)
    f_name = os.path.splitext(f_base)[0]

    target_folder = os.path.join(pdf_folder, f_name)
    if os.path.exists(target_folder):
        # if existing, remove it first
        shutil.rmtree(target_folder)

    # create folder first
    os.makedirs(target_folder)

    # Now begin to create new subpdfs
    inputpdf = PdfFileReader(open(f_path, "rb"))

    pAE = re.compile(settings.PDF_FILE_NAME_RE)
    new_pdf = True
    count = 1
    for i in range(inputpdf.numPages):
        if new_pdf:
            output = PdfFileWriter()
            
        page = inputpdf.getPage(i)
        output.addPage(page)

        content = convert_page_to_text(page)
            
        if split_key_word in content:
            saved_file_name = '{0}.pdf'.format(pAE.search(content).group(1))
            saved_file_path = os.path.join(target_folder, saved_file_name)
            with open(saved_file_path, 'wb') as f:
                output.write(f)
            new_pdf = True
            count += 1
        else:
            new_pdf = False
            

def get_email_list_from_excel():
    '''
    This function will read EXCEL FILE and get list of (email_address, name+email_address)
    Only works for xlsx and xlsm file format.
    '''

    f_path = settings.EXCEL_FILE
    
    wb = load_workbook(f_path)
    first_sheet_name = wb.get_sheet_names()[0]
    sheet = wb[first_sheet_name]

    title = []
    content = []
    first_row = True
    for row in sheet.iter_rows():
        row_value = []
        for cell in row:
            if first_row:
                title.append(cell.value)
            else:
                row_value.append(cell.value)

        if row_value:
            content.append(row_value)
        first_row = False

    email_set = set()

    name_key = settings.EXCEL_FILE_NAME_KEY
    email_keys = [v for v in title if v and v.lower().startswith('email')]
        
    for each_row in content:
        d = dict(zip(title, each_row))
        name = d[name_key]
        if not name:
            continue

        for email_key in email_keys:
            email_address = d[email_key]
            if not email_address:
                continue
            email_set.add((email_address, u'{0}<{1}>'.format(name, email_address)))

    email_list = list(email_set)
    email_list.sort(key=lambda x: x[0])
    return email_list

def get_split_pdf_list():

    pdf_list = []

    for subfolder in os.listdir(settings.PDF_FOLDER):
        fpath = os.path.join(settings.PDF_FOLDER, subfolder)
        if os.path.isdir(fpath):
            folder_file_list = []
            for fname in os.listdir(fpath):
                file_path = os.path.join(fpath, fname)
                if os.path.isfile(file_path):
                    folder_file_list.append(
                        (file_path, fname)
                    )

            if folder_file_list:
                pdf_list.append(
                    (subfolder, folder_file_list)
                )

    return pdf_list


def send_emails(subject, content, to_emails, pdf_file_list):
    email = EmailMessage(subject, content, settings.EMAIL_HOST_USER, to_emails)
    for pdf_file in pdf_file_list:
        email.attach_file(pdf_file, 'application/pdf')
    email.send()
