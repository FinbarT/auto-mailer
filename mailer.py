# -*- coding: utf-8 -*-
"""
Created on Thu Jun 15 11:25:12 2013
 
@author: Finbar Timmins
"""
import os
import ConfigParser
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email.header import Header
from email.Utils import formatdate
from email import Encoders
 
class customer(object):
    '''
    Customer object that contains name of customer, customers email address,
    path to their relevant attachments and the email itself. The email variable
    is only initialised at this stage.
    '''
    def __init__(self, config_dict):
        '''
        constructor takes a dictionary containing the config variables for the
        customer
        '''
        self.name = config_dict['name']
        self.email_address = config_dict['address']
        self.attachments_path = config_dict['attachments_path']
        self.mail = MIMEMultipart()
 
class mailer(object):
    '''
    Auto mailer object containing methods and variables for atuomatically
    sending emails to a large pool of customers
    '''
    def __init__(self):
        '''
        constructor that initialises an automailer object, consisting of a list
        of customers, message template, prompts the user for FedEx tracking
        number. Email username, password, and smpt server details come from a
        config file called smtp.ini
        '''
        self.mailing_list = []
        self.message_template = self.get_msg()
        self.package_tracking_num = raw_input(
            "Please enter the package tracking number: \n"
        )
        self.username = raw_input(
            "Please enter your email username:\n"
        )
        self.password = raw_input(
            "Please enter your email password:\n"
        )
        self.smtpserver = smtplib.SMTP()
        self.get_smtp_config()
         
    def get_smtp_config(self):
        '''
        Parses the config file smtp.ini to get the host and prot for the smtp
        connection
        '''
        config = ConfigParser.ConfigParser()
        config.read('smtp.ini')
         
        for section in config.sections():
            smtp_config = {}
            for option in config.options(section):
                smtp_config[option] = config.get(section, option)
 
        self.smtp_host = smtp_config['smtp_host']
        self.smtp_port = smtp_config['smtp_port']
 
    def build_customer_mailing_list(self):
        '''
        parses the customers details from the customers.ini config file. Then
        creates an instance the class customer for each and everyone with its
        associated name, email address and path to attachments
        '''
        config = ConfigParser.ConfigParser()
        config.read('customers.ini')
 
        for section in config.sections():
 
            customer_config = {}
            for option in config.options(section):
                customer_config[option] = config.get(section, option)
 
            self.mailing_list.append(customer(customer_config))
         
    def get_msg(self):
        '''
        reads in a dn returns the email message template
        '''
        msg = open('message_template.txt', 'rU')
        return msg.read()
 
    def build_emails(self):
        '''
        creates and email for each customer from the message template and
        attaches any relevant .pdf files availavle for each customer
        '''
        msg = ("Tracking Num: " +
               self.package_tracking_num +
               "\n\n" +
               self.message_template
        )
 
        for customer in self.mailing_list:
             
            content = MIMEText(
                msg,
                _charset='utf-8'
            )
 
            customer.mail['Subject'] = Header(
                customer.name +
                " card shipment",
                'utf-8'
            )
 
            customer.mail['From'] = self.username
            customer.mail['To'] = customer.email_address
            customer.mail['Date'] = formatdate(localtime=True)
            customer.mail.attach(content)
 
            for file_ in os.listdir(customer.attachments_path):
                if file_.lower().endswith(".pdf"):
 
                    path = os.path.join(customer.attachments_path, file_)
                    attachment = MIMEBase('application', 'octet-stream')
                    package = open(path, 'rb')
                    attachment.set_payload(package.read())
                    package.close()
                    Encoders.encode_base64(attachment)
 
                    attachment.add_header(
                            'Content-Disposition',
                            'attachment; filename="%s"' %
                            os.path.basename(path)
                    )
 
                    customer.mail.attach(attachment)
 
    def send_mail(self):
        '''
        starts the smtp connection, turns on tls mode and sends each email to
        each customers on the mailing list
        '''
 
        self.smtpserver.connect(self.smtp_host, self.smtp_port)
        self.smtpserver.starttls()
        self.smtpserver.login(self.username, self.password)
 
        for customer in self.mailing_list:
 
            self.smtpserver.sendmail(
                customer.mail['From'],
                customer.mail['To'],
                customer.mail.as_string()
            )
 
        self.smtpserver.quit()
 
def main():
     
    auto_mailer = mailer()
    auto_mailer.build_customer_mailing_list()
    auto_mailer.build_emails()
    auto_mailer.send_mail()
 
if __name__ == '__main__':
    main()
