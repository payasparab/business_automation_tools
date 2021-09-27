'''
Title: utils.email 
Author: Payas Parab (NinePointTwo Capital LLC)
Description: Utils related to automated emails/reporting

Key Methods:
 - generate_outlook_email: draft/send Outlook email

'''


import win32com.client
from win32com.client import Dispatch, constants    
import os
import os.path
    
def generate_outlook_email(send_to, subject, open_draft=True, body=None, attachment_path=None, 
                           cc=None, bcc=None, save_draft=False):
    '''
    Name: generate_outlook_email
    Description: Send/draft an Outlook email from Python  
    
    args: 
     - send_to : str/list : recepient(s) 
     - subject : str : Subject line of email

    kwargs:
     - open_draft : bool : def:True : Set to False to Send automatically
     - save_draft: bool : def:False : Set to True to Save to Drafts Folder
     - body: str : body of email
     - attachment_path: WindowsPath/str/list : File/folder(level 1 files)/List of Filepaths 
        > WARNING: Does not capture subfolders
     - cc: str/list : emails to CC
     - bcc: str/list : emails to BCC
    

    '''
    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    
    newMail.Subject = subject
    
    # Receipients #
    def recepient_str_converter(list_or_str):
        if type(list_or_str)==list:
            recipient_str = ''
            for i in list_or_str:
                recipient_str += (i+';')
        elif type(list_or_str)==str:
            recipient_str = list_or_str
        else: 
            recipient_str = None
        return recipient_str
    
    _to = recepient_str_converter(send_to)
    _cc = recepient_str_converter(cc)
    _bcc = recepient_str_converter(bcc)
    newMail.To = _to
    if _cc:
        newMail.CC = _cc
    if _bcc:
        newMail.BCC = _bcc
            
    if body:
        newMail.HTMLBody = body
    
    # Attachments #
    def attach_files(newMail, file_or_folder):
        if os.path.isfile(file_or_folder):
            newMail.Attachments.Add(Source=file_or_folder)
        elif os.path.isdir(file_or_folder):
            ## Iterate through dir
            for f in os.listdir(file_or_folder):
                _f_path = file_or_folder + '\\' + f
                if os.path.isfile(_f_path):
                    newMail.Attachments.Add(Source=_f_path)
        else: 
            raise AssertionError('Invalid attachment_path: {}'.format(file_or_folder))

    if attachment_path:
        if type(attachment_path) == str:
            attach_files(newMail, attachment_path)
        elif type(attachment_path) == list:
            for attachment in attachment_path:
                attach_files(newMail, attachment)
        else:
            raise ValueError('Invalid attachment_path. Not str or list.')


    # Action #
    if save_draft:
        newMail.save()  
    elif open_draft:
        newMail.display(True)
    else:
        newMail.Send()
        