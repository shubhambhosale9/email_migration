import imaplib
import os
import sys
import csv
import logging
import imapcopy
import subprocess
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from azure.storage.queue import QueueService
import uuid
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

logger = logging.getLogger('')
logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)

def imapcopy_automate(data, account_name, account_key, queue_name):
    #import ipdb; ipdb.set_trace()
    logger.info("The Email Credentials As follows :- %s" % (data))
    unique_filename = data[0]
    log_file = open(""+unique_filename+".txt", 'a+')
    log_file.write("Started Migration from ==>> %s =to==>> %s \n" %(data[0], data[3]))
    log_file.write("Source Host Server:- %s\n" %(data[2]))
    log_file.write("Destination Host Server:- %s\n\n\n" %(data[5]))
    log_file.close()
    logger.info("Started Migration from ==>> %s =to==>> %s " %(data[0], data[3]))
    logger.info("Source Host Server:- %s" %(data[2]))
    logger.info("Destination Host Server:- %s" %(data[5]))
    source_com = data[0] + ":" + data[1]
    dest_com = data[3] + ":" + data[4]
    logger.info("Source Email :- %s" % (source_com))
    logger.info("Destination Email :- %s" % (dest_com))
    ## connect to host using SSL
    imap_host_port = data[2]
    host = imap_host_port.split(':')
    imap = imaplib.IMAP4_SSL(host[0])

    ## login to server
    imap.login(data[0], data[1])
    ## list folders
    status, folder_list = imap.list()
    pure_folder_list = []
    for key, value in enumerate(folder_list):
        folder = value.split('"'+' '+'')[-1]
        name = folder.strip('"')
        pure_folder_list.append(name)
    
    attach_string = ''
    for each_folder in pure_folder_list:
        value = '\n' + each_folder
        attach_string  += value
    
    #import ipdb; ipdb.set_trace()
    msg = MIMEMultipart()
    msg['From'] = 'shubham@buzinessware.com'
    msg['To'] = "shubham@buzinessware.com,shubham32764@gmail.com, kaushalya@buzinessware.com"
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = 'Attention Regarding Email Migration Process Start'
    message = 'Hello Team, \n\nMigration of Email Account form %s to %s Account gets started.\nThe folder names as follows:- %s' %(data[0],                  data[3],attach_string)
    msg.attach(MIMEText(message, 'plain'))
    server = smtplib.SMTP('smtp.outlook.com:587')
    server.starttls()
    server.login('shubham@buzinessware.com', 'Naman@123!@#')
    server.sendmail(msg['From'], msg['To'].split(","), str(msg))
    server.close()
    logger.info('migration start mail notification sent sucessfully ')

    #import ipdb; ipdb.set_trace()

    try:
        for index, x in enumerate(folder_list):
            #if index==0:
            #    continue;
	    #sor = x.split('"/"')[1][1:]
            sor = x.split('"'+' '+'')[-1]
            ln = sor.split('.')
	    if len(ln)>=2:
	        if 'Sent' in ln[1]:
                    dest = 'sent items'
                elif 'Trash' in ln[1]:
		    dest = 'Deleted Items'
                elif 'Junk' in ln[1]:
		    dest = 'Junk Email'
                elif 'spam' in ln[1]:
		    dest = 'Junk'
                else:
		    dest = ln[1]
            else:
	        dest = ln[0]

	    sor1 = sor.strip('"')
            dest1 = dest.strip('"')
            logger.info("The folder name which we copy from source to destination:- %s, %s"  %(sor1, dest1))
            output = os.system('python imapcopy.py \"%s\" \"%s\"  \"%s\" \"%s\"   \"%s\"   \"%s\"  -c --verbose ' % (data[2], source_com,                                    data[5], dest_com, sor1, dest1))

    except Exception as e:
        logger.info("Exception caought in Automation script:- %s" %(e))

    finally:
        #import ipdb; ipdb.set_trace()
        logger.info("storing a migration log data onto Azure cloud storage")
        try:
            with open(""+unique_filename+".txt", 'r') as data:
                data_form = data.read()
                queue_service = QueueService(account_name, account_key)
                queue_service.create_queue(queue_name)
                queue_service.put_message(queue_name,  u''+data_form+'')
                logger.info('Log data inserted into azure queue sucessfully')
        except Exception as e:
            logger.info("Exception:- %s" %(e))

    #import ipdb; ipdb.set_trace()
    msg = MIMEMultipart()
    msg['From'] = 'shubham@buzinessware.com'
    msg['To'] = "shubham@buzinessware.com,shubham32764@gmail.com,kaushalya@buzinessware.com"
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = 'Attention Regarding Email Migration Process End'
    msg.attach( MIMEText("Please  find the bellow attachments and confirm migration is done on account %s" %(unique_filename)))
    part = MIMEBase('application', "octet-stream")
    part.set_payload( open(""+unique_filename+".txt", "r").read())
    Encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename= %s' %(""+unique_filename+".txt"))
    msg.attach(part)
    server = smtplib.SMTP('smtp.outlook.com:587')
    server.starttls()
    server.login('shubham@buzinessware.com', 'Naman@123!@#')
    server.sendmail(msg['From'], msg['To'].split(","), msg.as_string())
    server.close()
    logger.info('Mail sent sucessfully after migration process')


if __name__ == "__main__":
    #import ipdb; ipdb.set_trace()
    try:
        exist = os.path.isfile('output.txt')
        if exist :
            os.system('rm output.txt')
    except Exception as e:
        pass

    len_argv = len(sys.argv)
    if len_argv == 7:
        data = sys.argv
        imapcopy_automate(data[1:], 'storageaccount32764', '7FLoUx0dHpePFk98c0tVD0PAPYvMyHB81SD9byr+WM2lwz5ru7m8ekkFisuJBlWV5kLyAi4E8mx5aBfyUw                          2ecg==', 'shubham32764')
    elif len_argv == 1:
        try:
            with open('account.csv', "rb") as f_obj:
                reader = csv.reader(f_obj)
                reader.next()
	        for row in reader:
                    data = row
	            imapcopy_automate(data, 'storageaccount32764', 'pPHEuaVCsNFxWx/9GdjaHwZ0b/eIaliYsuVzUFCqWgW9kt1kpOwvC13oM4Ij8hJ7eyGYgKIlw+                                      74hP0O1bX4nQ==', 'shubham32764')
        except Exception as e:
            logger.info("Exception caught in main Of Automation script:- " + str(e))
    else:
        logger.info("Please Enter the proper inputs")
        sys.exit(2)
