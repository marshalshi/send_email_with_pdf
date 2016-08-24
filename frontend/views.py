import os

from django.shortcuts import render
from django import forms
from django.conf import settings

from crispy_forms.helper import FormHelper
from crispy_forms.layout import Layout

from .utils import split_pdf, get_email_list_from_excel, get_split_pdf_list, send_emails

#################################################################
#
# Send email to users
#

class SendEmailForm(forms.Form):
    email_list = forms.MultipleChoiceField(required=True, help_text='Target emails')
    file_list = forms.MultipleChoiceField(required=True, help_text='Attach PDF(s)')
    subject = forms.CharField(initial=settings.DEFAULT_EMAIL_SUBJECT)
    email_content = forms.CharField(widget=forms.Textarea, initial=settings.DEFAULT_EMAIL_CONTENT)

    def __init__(self, *args, **kwargs):
        super(SendEmailForm, self).__init__(*args, **kwargs)
        self.helper = FormHelper()
        self.helper.form_tag = False
        self.helper.form_class = 'form-horizontal'
        self.helper.label_class = 'col-md-3'
        self.helper.field_class = 'col-md-7'

        self.fields['email_list'].choices = get_email_list_from_excel()
        self.fields['file_list'].choices = get_split_pdf_list()

def send_email(request):

    msg = ''
    if request.method == 'POST':
        form = SendEmailForm(request.POST)
        if form.is_valid():
            to_emails = form.cleaned_data['email_list']
            pdf_file_list = form.cleaned_data['file_list']
            subject = form.cleaned_data['subject']
            email_content = form.cleaned_data['email_content']
            send_emails(subject, email_content, to_emails, pdf_file_list)
            msg = 'Email sent.'
    else:
        form = SendEmailForm()
    
    return render(request, 'send_email.html', {
        'nav': 'send_email',
        'form': form,
        'msg': msg
    })


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
