from win32com.client import Dispatch
from win32com.client import pywintypes
import datetime
import tempfile
import smtplib, os
import email
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import re,sys
import msvcrt
from lepl.apps.rfc3696 import Email

notesServer  = 'NOTES'
notesFile    = 'mailjrn.nsf'
notesPass    = 'xxXX1234'
relay        = '10.60.0.1'
 
# Connect to notes database (returns NotesDatabase object)
try:
  notesSession = Dispatch('Lotus.NotesSession')
  notesSession.Initialize(notesPass)
  notesDatabase = notesSession.GetDatabase(notesServer, notesFile)
except pywintypes.com_error as ex:
  print(ex.strerror,ex.excepinfo[2])
  sys.exit()
validate_email = Email();

# Given a document, return a list of attachment filenames and their contents
def extractAttachments(document):
    # Prepare
    attachmentPacks = []
    # For each item,
    for whichItem in range(len(document.Items)):
        # Get item
        item = document.Items[whichItem]
        # If the item is an attachment,
        if item.Name == '$FILE':
            # Get the attachment
            fileName = item.Values[0]
            fileBase, separator, fileExtension = fileName.rpartition('.')
            attachment = document.GetAttachment(fileName)
            attachment.ExtractFile(fileName)
            attachmentPacks.append(fileName)
    return attachmentPacks

def ConvertLotusSender(sender):
    # Если к нам приходит email пропускаем его
    if validate_email(sender):
      return (sender)
    # Если приходит Notes ID преобразовываем его в email  
    matches = re.search('CN=(.*?)/O',sender)
    if sender:
      seq = matches.group(1).split(' ')
      return(".".join(seq)+"@"+notesServer+".local")
    
def ConvertLotusRecipients(rcpt):
    rcpt_list = []
    for r in rcpt:
      rcpt_list.append(ConvertLotusSender(r))
    return rcpt_list

def sendMail(fromWhom, to, subj, text, att):
  msg = MIMEMultipart('alternative')
  msg['Subject'] = str(subj)
  msg['From'] = str(fromWhom)

  msg['To'] = ",".join(to)

  part1 = MIMEText(text,'plain')
  msg.attach(part1)
  i=0
  for attach in att:
    f = att
    part = MIMEBase('application', "octet-stream")
    part.set_payload( open(f[i],"rb").read() )
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="{0}"'.format(f[i]))
    msg.attach(part)
    os.remove(f[i])
    i+=1
  
  try:
    s = smtplib.SMTP(relay)
    s.sendmail(fromWhom, to, msg.as_string())
    s.quit()
  except smtplib.SMTPException:
    print(smtplib.SMTPException)
  except ConnectionRefusedError:
    print("Нет связи с сервером ТМ: %s" % relay)
    sys.exit()
  
def PrepareAndSend(document):
  subject = document.GetItemValue('Subject')[0].strip()
  date = document.GetItemValue('PostedDate')[0]
  fromWhom = ConvertLotusSender(document.GetItemValue('From')[0].strip())
  toWhoms = ConvertLotusRecipients(document.GetItemValue('SendTo'))
  if document.GetItemValue('CopyTo')[0]:
    toWhoms += ConvertLotusRecipients(document.GetItemValue('CopyTo'))
  if document.GetItemValue('BlindCopyTo')[0]:
    toWhoms += ConvertLotusRecipients(document.GetItemValue('BlindCopyTo'))
  body = document.GetItemValue('Body')[0].strip()
  attachments = extractAttachments(document)
  sendMail(fromWhom, toWhoms, subject, body, attachments)
  


folder = notesDatabase.GetView('$All')
# while True: 
print("Press Ctrl+x to exit")
while True:
  document = folder.GetFirstDocument()
  if document:
    if document.GetItemValue('Form')[0] == 'Memo': # remove NDR
      PrepareAndSend(document)
    document.RemovePermanently(True)
  # Exit by Ctrl+x  
  if msvcrt.kbhit():
    if ord(msvcrt.getch()) == 24:
      break
    
