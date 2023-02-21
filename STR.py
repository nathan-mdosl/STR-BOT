# import required packages
import os, time
from ftplib import FTP
from dateutil import parser
from pathlib import Path
import pandas as pd
from datetime import date
import shutil
import simplejson as json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage

import smtplib
import paramiko
import pysftp
import xlsxwriter
import openpyxl

paramiko.util.log_to_file('debug.txt', level = 'DEBUG')

root = Path(__file__).parent.resolve()


class STRBOT:
    def __init__(self):
        print("Starting Bot")
        print("Deleting old files...")
        mydir = "C:\\Users\\Nathan\\Desktop\\Projects\\jira\\sftp\\Local\\"  # change
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

        #sftp credits
        with open(root / 'credentials.json', "r") as rf:
             decoded_data = json.load(rf)
             for p in decoded_data['SFTP']:
                 sftpHost=p['sftpHost']
                 sftpPort=p['sftpPort']
                 sftpUser=p['sftpUser']
                 sftpPassword=p['sftpPassword']
                
        self.host = sftpHost
        self.port = sftpPort
        self.user = sftpUser
        self.password = sftpPassword
        self.filePath = "list.csv"  # taken from root of working dir
        self.cnopts = pysftp.CnOpts()
        self.cnopts.hostkeys = None

        #ftp credits
        with open(root / 'credentials.json', "r") as rf:
             decoded_data = json.load(rf)
             for p in decoded_data['FTP']:
                 ftpHost=p['ftpHost']
                 ftpUser=p['ftpUser']
                 ftpPassword=p['ftpPassword']

        
        self.ftpHost = ftpHost
        self.ftpUser = ftpUser
        self.ftpPassword= ftpPassword
    
        self.df = pd.read_csv(self.filePath)  # use pandas to read the csv file
        # extract all from the csv file
        FileNames = self.df.FileName.to_numpy()
        HotelNames = self.df.HotelName.to_numpy()
        FTPPaths = self.df.FTPPath.to_numpy()
        # iterate over extracted  to process them
        for self.FileName, self.HotelName, self.FTPPath in zip(FileNames, HotelNames,FTPPaths):
            self.ReadFromCSV(str(self.FileName).strip(), str(self.HotelName).strip(),str(self.FTPPath).strip())
        print("#######################")
        print("Bot has completed running")

    def ReadFromCSV(self, FileName, HotelName,FTPPath):
        print("#######################")
        print("Starting process. Reading CSV")
        print("Loaded file name is: " ,self.FileName, self.HotelName)

        self.Sftp_conn(FileName,self.cnopts,self.user,self.password)
    
    def Sftp_conn(self,host, port, username, password):
        print("#######################")
        print("Attemping to connect to:", self.host)
        try:

            with pysftp.Connection(self.host, port=self.port, username=self.user, password=self.password, cnopts=self.cnopts) as sftp:

                print("Bot has connected successfully to:", self.host)
                sftp.cwd("/Outbound")
                print("Switched to current dir:", sftp.pwd)
                self.latest = 0
                self.latestfile = None
                print("Searching for newest File in", sftp.pwd+".....")

                for fileattr in sftp.listdir_attr():

                    if fileattr.filename.startswith(self.FileName) and fileattr.st_mtime > self.latest:
                        
                        self.latest = fileattr.st_mtime
                        self.latestfile = fileattr.filename

                if self.latestfile is not None:
                    print("Found latest file:", self.latestfile)
                    os.getcwd()
                    self.localDIR = os.path.join(root, f"Local\\{self.HotelName}\\")
                    Path(self.localDIR).mkdir(parents=True, exist_ok=True)
                    os.chdir(self.localDIR)
                    os.getcwd()
                    print("Downloading file to:", self.localDIR)
                    sftp.get(self.latestfile, localpath=os.path.join(self.localDIR, self.latestfile))
                    print("Download done")
                    self.CheckFileFTP()
                    time.sleep(3)

        except Exception as e:
            print("Error: ", e)

    def CheckFileFTP(self):
        print("#######################")
        print("Checking if File is already uploaded")
        try:

            self.ftp = FTP(self.ftpHost, self.ftpUser, self.ftpPassword, timeout=200)
            print("Successfully connected to FTP server:", self.ftpHost)
            print("Changing DIR")
            self.ftp.cwd(self.FTPPath)
            self.ftp.dir()
            names = self.ftp.nlst()
            final_names= [line for line in names if self.latestfile in line]
            latest_time = None
            latest_name = None

            for name in final_names:
                ftptime = self.ftp.sendcmd("MDTM " + name)
                if (latest_time is None) or (ftptime > latest_time):
                    latest_name = name
                    latest_time = ftptime
                    print('Latest file {} is here. No need to upload'.format(self.latestfile))
                    self.ftp.close()
                
                break
                    
            else:
                print(self.latestfile, "Is not available in DIR. Preparing to Edit file. Before uploading to:", self.ftpHost)
                self.ftp.close()
                self.EditFile()
                time.sleep(3)


        except Exception as e:
            print(e)

    def EditFile(self):
        print("#######################")
        print("Attempting to edit:", self.latestfile)
        try:
            df = pd.read_excel(self.latestfile)
            df1=df.reindex(columns= ['CensusID','ChainID','HotelName','DateTY','DateLY','PropSupTY','PropDemTY','PropRevTY','PropSupLY','PropDemLY','PropRevLY','CompSupTY','CompDemTY','CompRevTY','CompSupLY','CompDemLY','CompRevLY'])
            writer = pd.ExcelWriter(self.latestfile,engine = "xlsxwriter")
            workbook  = writer.book
            df1.to_excel(writer, index = False)
            print(df1)
            time.sleep(3)
            writer.close()

            print("Removing bold style from columns")
            df = pd.read_excel(self.latestfile)
            column_list = df.columns
            writer = pd.ExcelWriter(self.latestfile, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            for idx, val in enumerate(column_list):
                worksheet.write(0, idx, val)

            writer.close()
            print("Excel file has been updated. Preparing to upload.")
            self.UploadFtp()
            time.sleep(3)
        except Exception as e:

            print("Error: ", e)
                
    def UploadFtp(self):
        print("#######################")
        print("Starting upload process to FTP server:" ,self.ftpHost)
        time.sleep(2)
        try:

            self.ftp = FTP(self.ftpHost, self.ftpUser, self.ftpPassword, timeout=200)
            print("Successfully connected to FTP server:", self.ftpHost)
            print("Changing DIR")
            self.ftp.cwd(self.FTPPath)
            time.sleep(2)
            print("Current DIR:", self.ftp.pwd())
            try:

                with open(self.latestfile, "rb") as file:
                    print("Uploading file:", self.latestfile)
                    # use FTP's STOR command to upload the file
                    self.ftp.storbinary(f"STOR {self.latestfile}", file)
                    
                self.ftp.nlst()
                print("File Successfully upload to DIR:", self.ftp.pwd())
      
                subject = 'STR Bot Has finished Running'
                message =(f"""STR BOT has successfully ran
File name {self.latestfile} has been downloaded,edited and uploaded to server {self.ftpHost} 
                


[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]
                    """)    
             
                try: # send email from given account to the  mail from csv, saying that GB check failed
                            self.sendEmail(self.errorReportemail, self.errorReportemailpasswd, self.errorReportemail, subject, message)
                            for mail in self.ccmail:
                                self.sendEmail(self.errorReportemail, self.errorReportemailpasswd, mail, subject, message)
                            time.sleep(2)
                except Exception as e:
                            print('Error while Sending failure email \nError Details:', e)
                           
                self.ftp.close()        
            except Exception as e:
                print("Error: ", e)

        except Exception as e:
            print("Error: ", e)
            
    def sendEmail(self, email, password, send_to, subject, message):

        print("Sending email to", send_to)
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
  

def main():
    STRBOT()


if __name__ == "__main__":
    main()