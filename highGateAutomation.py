from email.mime import text
import os
from pathlib import Path
import logging
import schedule
import time
import json
from win32api import SetDllDirectory
import win32com.client as wincl
import pysftp

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date

import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase

import fnmatch

root = Path(__file__).parent.resolve()

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(message)s',
                    filename=os.path.join(root,"app.log"),
                    filemode='a')

class  HighGateAutomation():
    def __init__(self):
        """ Initiating HighGate Automation Process Reading json details"""
        print('starting mdohistorical data excel automaion ')
        try:

            with open(os.path.join(root,'botDetails.json'),'r') as file:
                data = json.load(file)

                for x in data['bot']:
                    self.errorReportemail = x['erroremail']
                    self.errorReportemailpasswd = x['errorpwd']
                    self.aocEmail = x['AOCEmail'] 
                    self.aocPwd = x ['AOCpwd']
                    self.aocError1 = x ['AOCError1']
                    self.aocError2 = x ['AOCError2']
                    

                with open(os.path.join(root,'excelFile.json'),'r') as file:
                    excelDetails  = json.load(file)
                    for y in excelDetails['excel']:
                        reportname = y['reportname']
                        timesleep = y['timesleep']
                        excelname =  y['excelname']
                        macroname = y['macroname']
                        recivername  = y['receivername']
                        reciveremail = y['receiveremail']
                        print(reportname)
                        print(recivername)


                    

                        self.runMacro(excelname,macroname)
                        time.sleep(timesleep)
                    # copying genreated files to s3 buckert
                        # self.copyToSFTP(reportname)
                        # should i time sleep after copying before sending email
                        time.sleep(timesleep)
                    # sending txt file attachement to david
                        #  def sendEmailWithAttachement(self,email,password,reportname,recivername,reciveremail):
                        report_txt = self.writingReportNames(reportname)

                        self.sendEmailWithAttachement(self.aocEmail,self.aocPwd,reportname,recivername,reciveremail,report_txt)


        except Exception as e:
            print(e)
            logging.info("error reading excel details   :{}".format(e))
        
            subject = 'HighGate Automation Error reading json:'.format(e)
            message =(f"""HighGate Automation Error in reading json:
check above error
[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]
            """)
            try:
                self.sendErrorEmail(self.aocEmail,self.aocPwd,self.aocError1,self.aocError2,subject,message)
            except Exception as e:
                print('Error while Sending failure email \nError Details: {}'.format(e))
                logging.warning('Error while Sending QC email \nError Details: {}'.format(e))


    def runMacro(self,excelname,macroname):
        """Executing Macro Excel"""
        file_path = os.path.join(root,excelname)
        if os.path.exists(file_path):
            excel_macro = wincl.DispatchEx("Excel.application")

            workbook = excel_macro.Workbooks.Open(
                Filename = file_path,ReadOnly  =1
            )
            print(12)
            try:
                print(14)
                # "basCollectData.GenerateAll"
                excel_macro.Application.Run(macroname)
                time.sleep(1)
                # pyautogui.click(930,624)
                workbook.Save()
                # pass
                print(21)
            except Exception as e:
                print(e)
                logging.info("error reading excel details   :{}".format(e))
           
                subject = 'HighGate Automation Error executing macro:'.excelname
                message =(f"""HighGate Automation Error executing macro: {e}
check above error
[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]
            """)
                try:
                    self.sendErrorEmail(self.aocEmail,self.aocPwd,self.aocError1,self.aocError2,subject,message)
                except Exception as e:
                    print('Error while Sending failure email \nError Details: {}'.format(e))
                    logging.warning('Error while Sending QC email \nError Details: {}'.format(e))

            
            finally:
                excel_macro.Application.Quit()
                


    def writingReportNames(self,reportname):
        """ Writing Report Names to txt file"""
        today = date.today()
        current_date= str(today)
        
        report  ="{}{}.txt".format(reportname,current_date)

        with open(report, "w") as a:

            reportlist  = (fnmatch.filter(os.listdir(root), '*.xlsb'))
            count = 1
            for filename in reportlist:
                row = "{}. {} {}".format(count,filename,os.linesep)
                a.write(row)
                count = count +1
            return report

    def sendEmailWithAttachement(self,email,password,reportname,recivername,reciveremail,report_txt):
        """ Sending Mail With txt attachement """
        
        report_count = len(fnmatch.filter(os.listdir(root), '*.xlsb'))
  
        today = date.today()
        subject = "Report Generate For {} On {}".format(reportname,today)



        body = """Hi {},

Report Type :{}
Date Generated:{}
Report count:{}

Please Find attached txt file containing reports names archived
Thank you.
[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL] 
"""    .format(recivername,reportname,today,report_count)     
        message = MIMEMultipart()
        message["From"] = email
        message["To"] = reciveremail
        message["Subject"] = subject

        message.attach(MIMEText(body, "plain"))
        str_today = str(today)
     
        filename =  report_txt
   

        with open(filename,"rb") as attachment:
            part  = MIMEBase("application","octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header(
           "Content-Disposition",
           f"attachement; filename= {filename}",
        )
        message.attach(part)
        text = message.as_string()

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com",465,context=context) as server:
            server.login(email,password)
            server.sendmail(email,reciveremail,text)
            print(body)
            print('send email')






    def copyToSFTP(self,reportname):
        """ Copying reports to sftp location """
        
        today = date.today()
        current_date = str(today)
        src = "{}{}".format(reportname,current_date)

        print(src)
        cnopts = pysftp.CnOpts()
        cnopts.hostkeys = None
        try:

            with pysftp.Connection('mypsftp.mydigitaloffice.com', username='hiton_GRO_ROOT', password='bLp5rTDjXWfqbVhK', cnopts=cnopts) as sftp:
                print(1)
                with sftp.cd("/mdo-hilton-gro-system/testing"):
                    localpath ="E:/python/highGate/macro"
                    remotepath = '/mdo-hilton-gro-system/testing/test/{}/'.format(src)

                   
                    sftp.put_d(localpath,remotepath,True)

                    directory_structure = sftp.listdir_attr()
                    print(directory_structure)
# printing folder and file structure of sftp location
                    # for attr in directory_structure:

                    #     print (attr.filename)
        except Exception as e:
            logging.info("error archiving  excel :{}".format(e))
           
            subject = 'error archiving  excel:'.excelname
            message =(f"""HighGate Automation Error archiving  files: {e}
check above error
[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]
            """)
            try:
                self.sendErrorEmail(self.aocEmail,self.aocPwd,self.aocError1,self.aocError2,subject,message)
            except Exception as e:
                print('Error while Sending failure email \nError Details: {}'.format(e))
                logging.warning('Error while Sending QC email \nError Details: {}'.format(e))

    def sendErrorEmail(self, email, password, send_to_1,send_to_2 ,subject, message):
        """Sending Error Email """
        print("Sending email to", send_to_1)
        logging.info("Sending error email")
        try:
            
            msg = MIMEMultipart()
            # recipients = [send_to_1,send_to_2]
            msg["From"] = email
            # msg["To"] = ", ".join(recipients)
            msg["To"] = send_to_1
            msg["Cc"] = send_to_2
        
            msg["Subject"] = subject
            msg.attach(MIMEText(message, 'plain'))
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(email, password)
            text = msg.as_string()
            server.sendmail(email, send_to_1, text)
            server.quit()
        except Exception as e:
            print('Error while Sending failure email \nError Details: {}'.format(e))
            logging.warning('Error while Sending failure email \nError Details: {}'.format(e))




def main():
    print(175)
    HighGateAutomation()

def AutomationSchedule(scheduledate):
    current_day = date.today().day
    print(type(current_day))

    if(current_day != scheduledate):
        return
    

    main()
    

    




if __name__ =='__main__':
    with open(os.path.join(root,'records.json')) as json_file:
        print(36)

        data = json.load(json_file)
        scheduleTime=data['sheduledTime']
        print(scheduleTime)
        isschedule = data['issheduled']
        print(scheduleTime)
        print(isschedule)
        schedule_date = data['scheduledate']

        if(isschedule == "False"):
            print(191)
            AutomationSchedule(schedule_date)
        else:
            
            schedule.every(1).day.at(scheduleTime).do(AutomationSchedule(schedule_date))
       

        while True:
            schedule.run_pending()
            time.sleep(1)

        


