import numpy as np
import pandas as pd

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from os.path import basename
import email
import email.mime.application

from fpdf import FPDF

dataset = pd.read_excel('student.xlsx')


for i in range(len(dataset)):
    
    #making a PDF file
    
    id1 = str(dataset['student_id'][i])
    name = dataset['name'][i] 
    email1 = dataset['email'][i] 
    marks = str(dataset['score'][i])
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Dear "+name+" (ID: "+id1+")"+",", ln=1, align="L")
    pdf.cell(200, 10, txt="You recieved "+marks+" marks in the exam.", ln=1, align="L")
    pdf.cell(200, 20, txt="Regards,", ln=1, align="L")
    pdf.cell(200, 5, txt="LEARNTRICKS", ln=1, align="L")
    pdf.output(name+"_result.pdf")
     
    #MAIL BODY
    
    html = """
     
    Dear Student,<br><br>
     
    Please Find Attached.<br><br> 
     
     
    Best Regards,<br>
    LEARNTRICKS
    """
     
    # Creating message.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Result"
    msg['From'] = "amitesh.srivastava11@gmail.com"
    msg['To'] = email1
     
    # The MIME types for text/html
    HTML_Contents = MIMEText(html, 'html')
    
    filename = name+"_result.pdf"
    fo=open(filename,'rb')
    attach = email.mime.application.MIMEApplication(fo.read(),_subtype="pdf")
    fo.close()
    attach.add_header('Content-Disposition','attachment',filename=filename)
     
    # Attachment and HTML to body message.
    msg.attach(attach)
    msg.attach(HTML_Contents)
    
    # Your SMTP server information
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('amitesh.srivastava11@gmail.com','password')  #EMAIL PASSOWRD HERE
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()