# import required packages
import os, time
import sys
from ftplib import FTP
import ftplib

from dateutil import parser
from pathlib import Path
import pandas as pd
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

import logging
import paramiko
import pysftp
import xlsxwriter

import datetime

ftpHost = "52.72.14.3"
ftpUser = "MYPUpload"
ftpPassword= "gB8m8DAQqz"

uploadToDir = "Highgate_week_"
print("Starting upload process to FTP server:" ,ftpHost)
time.sleep(2)


ftp = FTP(ftpHost,ftpUser,ftpPassword, timeout=200)
print("Successfully connected to FTP server:", ftpHost)
print("Changing DIR")
ftp.cwd("/Data/2479/958/STR/STRReport")
time.sleep(2)
print("Current DIR:", "/Data/2479/958/STR/STRReport")

# today = str(datetime.datetime.now().isoformat(sep="-",timespec="seconds"))

# s = today
# s = s.replace("/", "_").replace(":", "-")
# ftp.mkd(s)
# ftp.cwd(s)
ftp.dir()

names = ftp.nlst()
final_names= [line for line in names if uploadToDir in line]

latest_time = None
latest_name = None

for name in final_names:
    time = ftp.sendcmd("MDTM " + name)
    if (latest_time is None) or (time > latest_time):
        latest_name = name
        latest_time = time

        print("File is here")
    break
        
    
    
else:
    print("fail")





