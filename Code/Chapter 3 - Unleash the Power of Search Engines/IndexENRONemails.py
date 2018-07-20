emailFolder = "D:/enron_mail_20150507/maildir"
indexFolder = "D:/enron_mail_index"

## Get list of all files
from pathlib import Path
allDocs = Path(emailFolder).glob('**/*')
files = [x for x in allDocs if x.is_file()]
print("Number of files: ", len(files))

## Import libraries I will need
import os, os.path
from whoosh import fields, index
from whoosh.analysis import StemmingAnalyzer
import datetime
import email

## Define index schema
mySchema = fields.Schema(To = fields.STORED,
                         From = fields.STORED,
                         Subject = fields.TEXT(field_boost=2.0, stored=True),
                         Date = fields.STORED,
                         Message = fields.TEXT(analyzer=StemmingAnalyzer()),
                         File = fields.STORED)

## Start indexing
if not os.path.exists(indexFolder):
    os.mkdir(indexFolder)
ix = index.create_in(indexFolder, mySchema)
writer = ix.writer()

## Print start time
print(datetime.datetime.now())

## Loop through each email and index
for eachFile in files:
    myFile = open(eachFile, "r")
    emailData = myFile.read()

    tempEmail = email.message_from_string(emailData)
    tempSender = tempEmail['from']
    tempReceiver = tempEmail['to']
    tempSubject = tempEmail['subject']
    tempDate = tempEmail['date']

    if tempEmail.is_multipart():
        for part in tempEmail.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
        
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                tempMessage = part.get_payload(decode=True)
                break
    else:
        tempMessage = tempEmail.get_payload(decode=True)

    tempMessage = tempMessage.decode().replace('\\n', "\n")

    ## Only index if it is an email
    if (tempSender == ''):
        break
    
    ## Add email to index
    writer.add_document(To = tempReceiver,
                        From = tempSender,
                        Subject = tempSubject,
                        Date = tempDate,
                        Message = tempMessage,
                        File = eachFile)

writer.commit()

## Print end time
print(datetime.datetime.now())
