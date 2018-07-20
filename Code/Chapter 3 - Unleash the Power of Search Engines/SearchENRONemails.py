
searchlist = ['ESOP', '"Retirement Plan"', '"Savings Plan"',
              'Tax', 'Audit', 'Project Steele', 'Raptor',
              'LJM', 'Condor', 'JEDI', 'Dynergy', 'CalPERS',
              'Cornhusker', 'Slapshot', 'Shred', 'Fraud',
              'bankruptcy OR bankrupt', 'resignation OR resign',
              'bribe OR bribery', 'hide', 'Accounting AND Tax']

indexFolder = "D:/enron_mail_index"

from whoosh import index
from whoosh.qparser import MultifieldParser
import pandas as pd
import numpy as np

ix = index.open_dir(indexFolder)

for search in searchlist:
    myParser = MultifieldParser(["Subject", "Message"], 
                                schema = ix.schema)
    
    query = myParser.parse(search)
    
    with ix.searcher() as searcher:
        
        myEmails = searcher.search(query, limit = None)
       
        ## Print number of emails found
        print("Search String: ", search)
        print("Number of emails found: ", len(myEmails))
      
        ## Create a Pandas DataFrame to save results
        saveColumns = ["To", "From", "Subject", "Date", "File"]
        myIndex = np.arange(len(myEmails))
        data = pd.DataFrame(columns = saveColumns, index = myIndex)
        
        ## Store results to a file
        count = 0
        for x in myEmails:
            tempReceiver = ""
            tempSender = ""
            tempSubject = ""
            tempDate = ""
            
            if "To" in x.keys():
                tempReceiver = str(x["To"])
            if "From" in x.keys():
                tempSender = str(x["From"])
            if "Subject" in x.keys():
                tempSubject = str(x["Subject"])
            if "Date" in x.keys():
                tempDate = str(x["Date"])
            
            tempFile = str(x["File"])
            
            data.loc[count] = pd.Series({
                    'To': tempReceiver,
                    'From': tempSender,
                    'Subject': tempSubject,
                    'Date': tempDate,
                    'File': tempFile})
            count = count + 1
        
        filename = search.replace('"', '') + ".xlsx"
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        data.to_excel(writer, sheet_name="email list", index=False)
        writer.close()


