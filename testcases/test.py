import os, time
import sys
from ftplib import FTP
import ftplib
import logging
from dateutil import parser
from pathlib import Path
import pandas as pd
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from datetime import date
import shutil

import simplejson as json

from apscheduler.schedulers.blocking import BlockingScheduler

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage

import smtplib

root = Path(__file__).parent.resolve()


class FTPCON:
    def __init__(self):
        print("Starting Bot")
        mydir = "D:\\Nathan\\GSS s3\\Archive"  # change
        try:
            shutil.rmtree(mydir)
            print("Folder Successfully Deleted")
        except OSError as e:
            print("Error: %s - ." % (e.filename))
        
        with open(os.path.join(root, 'CC_Mails.json'), 'r') as file: #enter email and password for error reporting gmail account
            data = json.load(file)
            self.ccmail = data['ccmails']
            errormaildets = data['erroremail']
            self.errorReportemail = errormaildets['email']
            self.errorReportemailpasswd = errormaildets['password']
        
    
        self.host = "54.86.242.112"
        self.user = "mypuser"
        self.password = "myp@123"
        self.filePath = "D:\\Nathan\\GSS s3\\list.csv"  # change
        time.sleep(2)
        self.CreateGDriveFolder()
        
    def sendEmail(self, email, password, send_to, subject, message):

        print("Sending email to", send_to)
        logging.info("Sending email to")
        msg = MIMEMultipart()
        msg["From"] = email
        msg["To"] = send_to
        msg["Subject"] = subject
        msg.attach(MIMEText(message, 'plain'))
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, send_to, text)
        server.quit()

    def CreateGDriveFolder(self):
        today = str(date.today())
        print("Todays Date: " + today)

        gauth = GoogleAuth()
        self.drive = GoogleDrive(gauth)

        folder_name = today
        folder = self.drive.CreateFile(
            {
                "title": folder_name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [{"id": "1k1kFXJa1MTlAQrkPE0uKW-kO3Hj1Bjp5"}],
            }
        )
        time.sleep(2)
        print("Folder Created")
        folder.Upload()

        # Get folder info and print to screen
        self.foldertitle = folder["title"]
        self.folderid = folder["id"]
        print("title: %s, id: %s" % (self.foldertitle, self.folderid))
        NewestFolderID = self.folderid
        print("Folder ID :", NewestFolderID)
        self.SetupBOT(self.filePath)

    def SetupBOT(self, file):
        self.df = pd.read_csv(file)  # use pandas to read the csv file
        # extract all from the csv file
        self.HotelIDS = self.df.HotelID.to_numpy()
        HotelNames = self.df.HotelName.to_numpy()
        # iterate over extracted  to process them
        for self.HotelID, self.HotelName in zip(self.HotelIDS, HotelNames):
            self.ReadFromCSV(str(self.HotelID), self.HotelName.strip())

        print("Finished Downloading and Uploading Reports")
        subject = 'GSS Report Bot Has finished Running'
        message =(f"""Bot For GSS HighGate Select has successfully ran
Google Sheet will be ready by Today at 4 PM SL Time

Next run time will be next Wednesday
[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]
            """)    
             
        try: # send email from given account to the  mail from csv, saying that GB check failed
                    self.sendEmail(self.errorReportemail, self.errorReportemailpasswd, self.errorReportemail, subject, message)
                    for mail in self.ccmail:
                        self.sendEmail(self.errorReportemail, self.errorReportemailpasswd, mail, subject, message)
                    time.sleep(2)
        except Exception as e:
                    print('Error while Sending failure email \nError Details:')
                    logging.warning('Error while Sending QC email \nError Details:') 
        print("Deleting local folder before downloading reports")
        mydir = "D:\\Nathan\\GSS s3\\Archive"  # change
        try:
            shutil.rmtree(mydir)
            print("Folder Successfully Deleted")
        except OSError as e:
            print("Error: %s - ." % (e.filename))
        print("Next run time next Wednesday")

    def ReadFromCSV(self, HotelID, HotelName):
        print("Reading CSV")
        print(self.HotelID, self.HotelName)

        self.FTPConnect(HotelID)

    def FTPConnect(self, HotelID):

        try:

            print("Connecting to FTP DIR")
            time.sleep(2)

            self.ftp = FTP(self.host, self.user, self.password, timeout=200)

            self.ftp.cwd("/Archive/2479/")
            print("Current DIR: /Archive/2479/")
            # try:
            #     self.files = self.ftp.nlst()
            # except ftplib.error_perm as resp:
            #     if str(resp) == "550 No files found":
            #         print ("No files in this directory")
            #     else:
            #         raise

            # for f in self.files:
            #     print (f)

            print("Changing DIR")
            self.ftp.cwd(f"/Archive/2479/{self.HotelID}/Medallia/MedalliaReport/")
            print(f"Current DIR: /Archive/2479/{self.HotelID}/Medallia/MedalliaReport/")
            self.DownloadFTP(self)

        except ftplib.all_errors as e:
            errorcode_string = str(e).split(None, 1)[0]
            print(e)
            logging.info(errorcode_string)
            print("Error. Hotel ID doesn't exist! Going to next hotel ID")
            time.sleep(10)

    def DownloadFTP(self, HotelID):

        while True:
            try:
                print("Loaded Hotel ID is:  ", self.HotelID)
                time.sleep(2)
                self.files = self.ftp.nlst()
                # for f in self.files:
                #     print (f)

                lines = []
                self.ftp.dir("", lines.append)

                latest_time = None
                latest_name = None

                for line in lines:
                    tokens = line.split(maxsplit=9)
                    time_str = tokens[5] + " " + tokens[6] + " " + tokens[7]
                    time_ftp = parser.parse(time_str)
                    if (latest_time is None) or (time_ftp > latest_time):
                        latest_name = tokens[8]
                        latest_time = time_ftp
                print("Latest file is: " + latest_name)
                time.sleep(4)
                print("Downloading: " + latest_name)

                os.getcwd()
                self.localDIR = os.path.join(root, f"Archive\\{self.HotelID}\\")
                Path(self.localDIR).mkdir(parents=True, exist_ok=True)
                os.chdir(self.localDIR)
                os.getcwd()
                print(self.localDIR)

                file = open(latest_name, "wb")
                self.ftp.retrbinary("RETR " + latest_name, file.write)
                file.close()
                print("Renaming file")

                newname = str(self.HotelID)
                oldfilename = os.path.splitext(latest_name)[1]
                os.rename(latest_name, newname + oldfilename)
                print("New file name is: " + newname + oldfilename)

                print("Download Done")
                self.UploadGDrive(self)
                break

            except ftplib.all_errors as e:
                errorcode_string = str(e).split(None, 1)[0]
                print(e)
                logging.info(errorcode_string)
                print("Error. Trying again.")
                time.sleep(5)

    def UploadGDrive(self, HotelID):
        self.ftp.quit()

        folder = self.folderid
        ArchiveReportsDir = self.localDIR

        for f in os.listdir(ArchiveReportsDir):
            filename = os.path.join(ArchiveReportsDir, f)
            print("Uploading file: " + filename)
            gfile = self.drive.CreateFile({"parents": [{"id": folder}], "title": f})

            gfile.SetContentFile(filename)
            print("Upload completed")
            gfile.Upload()

            for i in range(3, 0, -1):
                sys.stdout.write(str(i) + " ")
                sys.stdout.flush()
                time.sleep(1)


def main():
    FTPCON()


if __name__ == "__main__":
    main()
