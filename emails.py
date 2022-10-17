import imaplib
import os
import email
import sys
import json
from types import NoneType
import lxml
from bs4 import BeautifulSoup
import xlwt
from tempfile import TemporaryFile


class GMAIL_EXTRACTOR():

    def initializeVariables(self):
        self.usr = ""
        self.pwd = ""
        self.mail = object
        self.mailbox = ""
        self.mailCount = 0
        self.destFolder = ""
        self.data = []
        self.ids = []
        self.idsList = []

    def getLogin(self):
        self.usr = "mubinawaheed1@gmail.com"
        self.pwd = "qlju cgjo abxa xsce"


    def attemptLogin(self):
        self.mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        if self.mail.login(self.usr, self.pwd):
            print("\nLogon SUCCESSFUL")
	try:
		self.destFolder = input("\nPlease choose a destination folder ")
            	if not self.destFolder.endswith("/"): self.destFolder+="/"
            	return True
	except:
		return False
		
        else:
            print("\nLogon FAILED")
            return False

#     def checkIfUsersWantsToContinue(self):
#         print("\nWe have found "+str(self.mailCount)+" emails in the mailbox "+self.mailbox+".")
#         return True if input("Do you wish to continue extracting all the emails into "+self.destFolder+"? (y/N) ").lower().strip()[:1] == "y" else False       
        
    def selectMailbox(self):
        self.mailbox = input("\nPlease type the name of the mailbox you want to extract, e.g. Inbox: ")
        bin_count = self.mail.select(self.mailbox)[1]
        self.mailCount = int(bin_count[0].decode("utf-8"))
        return True if self.mailCount > 0 else False

    def searchThroughMailbox(self):
        type, self.data = self.mail.search(None, "ALL")
        self.ids = self.data[0]
        self.idsList = self.ids.split()

    def parseEmails(self):
        jsonOutput = {}
        subjectlist=[]
        tolist=[]
        fromlist=[]
        datelist=[]
        bodylist=[]
        count=0
        for anEmail in self.data[0].split():
            if count>1000:
                break
            type, self.data = self.mail.fetch(anEmail, '(UID RFC822)')
            # print(self.data)
            raw = self.data[0][1]
            try:
                raw_str = raw.decode("utf-8")
            except UnicodeDecodeError:
                try:
                    raw_str = raw.decode("ISO-8859-1") # ANSI support
                except UnicodeDecodeError:
                    try:
                        raw_str = raw.decode("ascii") # ASCII ?
                    except UnicodeDecodeError:
                        pass
						
            msg = email.message_from_string(raw_str)

            # clean_cc=BeautifulSoup(msg['cc'], 'lxml').text
            # jsonOutput['cc']=clean_cc
            # cc.append(jsonOutput['cc'])

            clean_subject=BeautifulSoup(msg['subject'], "lxml").text
            jsonOutput['subject'] = clean_subject
            subjectlist.append(jsonOutput['subject'])

            clean_from=BeautifulSoup(msg['from'], "lxml").text
            jsonOutput['from'] = clean_from
            # print("from",clean_from)
            fromlist.append(jsonOutput['from'])

            clean_date=BeautifulSoup(msg['date'], "lxml").text
            jsonOutput['date'] = clean_date
            datelist.append(jsonOutput['date'])

            if msg['to'] is not NoneType:
                
                clean_to=BeautifulSoup(msg['to'], "lxml").text
                jsonOutput['to']=clean_to
                tolist.append(jsonOutput['to'])
                # print(tolist)

            else:
                print(msg['to'])
                continue

            # print("-----------------to--------\n",jsonOutput['to'])
            
            raw = self.data[0][0]
            raw_str = raw.decode("utf-8")
            uid = raw_str.split()[2]
            # if int(uid)>100: break
            # Body #
            if msg.is_multipart():
                for part in msg.walk():
                    partType = part.get_content_type()
                    ## Get Body ##
                    if partType == "text/plain" and "attachment" not in part:
                        body = part.get_payload()
                        clean_body=BeautifulSoup(body, "lxml").text
                        jsonOutput['body']=clean_body
                        count+=1

                    # print("\n-----------------------------------\n",jsonOutput['body'])
                    ## Get Attachments ##
                    # if part.get('Content-Disposition') is None:
                    #     attchName = part.get_filename()
                    #     if bool(attchName):
                    #         attchFilePath = str(self.destFolder)+str(uid)+str("/")+str(attchName)
                    #         os.makedirs(os.path.dirname(attchFilePath), exist_ok=True)
                    #         with open(attchFilePath, "wb") as f:
                    #             f.write(part.get_payload(decode=True))
            else:
                body = msg.get_payload(decode=True).decode("unicode-escape")
                clean_body=BeautifulSoup(body, "lxml").text
                jsonOutput['body']=clean_body
                # print("\n------------body",jsonOutput['body'])
                bodylist.append(jsonOutput['body'])
                print(jsonOutput['body'])
                count+=1
                
            # outputDump = json.dumps(jsonOutput)
            # emailInfoFilePath = str(self.destFolder)+str(uid)+str("/")+str(uid)+str(".csv")
            # os.makedirs(os.path.dirname(emailInfoFilePath), exist_ok=True)
            # with open(emailInfoFilePath, "w") as f:
            #     f.write(outputDump)

            
        # print(subjectlist, bodylist, tolist, datelist, fromlist)

        book = xlwt.Workbook()
        sheet1 = book.add_sheet('sheet1')

        for i,e in enumerate(subjectlist):
            e = e.replace("\r\n","")
            e=e.replace("=?utf-8?q", "")
            sheet1.write(i+1,0,e)

        for i,e in enumerate(fromlist):
            e = e.replace("\r\n","")

            sheet1.write(i+1,1,e)

        for i,e in enumerate(datelist):

            e = e.replace("\r\n","")
            sheet1.write(i+1,2,e)

        for i,e in enumerate(tolist):
            e = e.replace("\r\n","")
            sheet1.write(i+1,3,e)

        for i,e in enumerate(bodylist):
            e = e.replace("\r\n","")
            sheet1.write(i+1,4,e)

        name = "random.xls"
        book.save(name)
        book.save(TemporaryFile())
        print('doone')
        print(subjectlist)

        # with open("myemails.csv", "a") as f:
                # f.write(outputDump)
        

    def __init__(self):
        self.initializeVariables()
        self.getLogin()
        if self.attemptLogin():
            not self.selectMailbox() and sys.exit()
        else:
            sys.exit()
#         not self.checkIfUsersWantsToContinue() and sys.exit()
        self.searchThroughMailbox()
        self.parseEmails()

if __name__ == "__main__":
    run = GMAIL_EXTRACTOR()



