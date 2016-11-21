# -*- coding: utf-8 -*-
import os
import json
from collections import defaultdict

from django.shortcuts import render
from django import forms
from django.conf import settings
from django.http import JsonResponse, HttpResponseNotFound

from crispy_forms.helper import FormHelper
from crispy_forms.layout import Layout

from .utils import split_pdf, get_excel_content, get_split_pdf_dict, send_emails

#################################################################
#
# Send email to users
#

def send_email(request):

    excel_content = get_excel_content()
    pdf_dict = get_split_pdf_dict()
    
    return render(request, 'send_email.html', {
        'nav': 'send_email',
        'excel_content': excel_content,
        'pdf_dict': pdf_dict,
        'DEFAULT_EMAIL_SUBJECT': settings.DEFAULT_EMAIL_SUBJECT,
        'DEFAULT_EMAIL_CONTENT': settings.DEFAULT_EMAIL_CONTENT
    })

def send_email_json(request):

    if request.method == 'POST':
        subject = request.POST.get('subject', settings.DEFAULT_EMAIL_SUBJECT)
        content = request.POST.get('content', settings.DEFAULT_EMAIL_CONTENT)
        send_type = request.POST.get('send_type', 'single')
        file_email_map = json.loads(request.POST.get('file_email_map', "{}"))

        if send_type == 'bundle':
            
            email_file_map = defaultdict(list)
            for pdf_file, to_emails in file_email_map.iteritems():
                to_emails = [em for em in to_emails if em]
                to_emails.sort()
                email_file_map[tuple(to_emails)].append(pdf_file)

            for to_emails, pdf_file_list in email_file_map.iteritems():
                send_emails(subject, content, to_emails, pdf_file_list)

        else:
            # default send mode is single
            for pdf_file, to_emails in file_email_map.iteritems():
                to_emails = [em for em in to_emails if em]
                send_emails(subject, content, to_emails, [pdf_file,])
        
    else:
        return HttpResponseNotFound()

    return JsonResponse({'success': True})



#################################################################
#
# Upload PDF and split it to sub pdfs
#

class PDFUploadForm(forms.Form):
    pdf_file = forms.FileField(
        label='PDF file',
        help_text='The PDF file which need to split.',
        required=True
    )

    def __init__(self, *args, **kwargs):
        super(PDFUploadForm, self).__init__(*args, **kwargs)
        self.helper = FormHelper()
        self.helper.form_tag = False
        self.helper.form_class = 'form-horizontal'
        self.helper.label_class = 'col-md-4'
        self.helper.field_class = 'col-md-8'        
    
    def clean(self, *args, **kwargs):
        data = super(PDFUploadForm, self).clean(*args, **kwargs)
        file = data.get('pdf_file')
        if file:
            # check whether the file is pdf
            # if no raise error
            filename = file.name
            ext = os.path.splitext(filename)[1]
            ext = ext.lower()
            if ext not in ['.pdf']:
                raise forms.ValidationError("Not allowed file type.")
        return data
        

def handle_uploaded_file(f, fpath):
    # save file to dist
    with open(fpath, 'wb+') as destination:
        for chunk in f.chunks():
            destination.write(chunk)

def upload_pdf(request):

    msg = ''
    if request.method == 'POST':
        form = PDFUploadForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = form.cleaned_data['pdf_file']
            filename = uploaded_file.name
            fpath = os.path.join(settings.MEDIA_ROOT, filename)
            handle_uploaded_file(uploaded_file, fpath)
            split_pdf(fpath, settings.PDF_SPLIT_KEY_WORDS)
            msg = 'File uploaded and splited.'
    else:
        form = PDFUploadForm()

    return render(request, 'upload_pdf.html', {
        'nav': 'upload_pdf',
        'form': form,
        'msg': msg
    })
