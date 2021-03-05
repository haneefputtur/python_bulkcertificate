#Script developed by haneef puttur to bulk generate certificates in pdf and send email
from reportlab.lib.pagesizes import A4
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
import pandas as pd
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 


def sendmail(name, toaddr, attach_file):
    fromaddr = "fasfasfasf@gmail.com"
    apppassword = "REPLACE APP PASSWORD" 
    
    msg = MIMEMultipart() 
  
    # storing the senders email address   
    msg['From'] = fromaddr 
  
    # storing the receivers email address  
    msg['To'] = toaddr 
  
    # storing the subject  
    msg['Subject'] = "Certificate from IEEE"
  
    # string to store the body of the mail 
    body = "Dear "+name+", <br><br>Please find attached your certificate. <br><br> Best Regards <br><br>Haneef <br><br>For Digix Online Solutions."
  
    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'html')) 
  
    # open the file to be sent  
    filename = attach_file.split('//')[1]
    attachment = open(attach_file, "rb") 
    
  
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
  
    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 
  
    # encode into base64 
    encoders.encode_base64(p) 
   
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
  
    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) 
    
    s = smtplib.SMTP('smtp.gmail.com', 587) 
  
    # start TLS for security 
    s.starttls() 
  
    # Authentication 
    s.login(fromaddr, apppassword) 
  
    # Converts the Multipart msg into a string 
    text = msg.as_string() 
  
    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 
  
    # terminating the session 
    s.quit() 
    print('\nEmail Sent : '+toaddr)
    attachment.close()


def make_certificate(name, email):
    file_name = name.replace(" ","_")
    file_name = 'pdf//'+file_name+'.pdf'
    img = 'cert_template.jpg'
    c = canvas.Canvas(file_name, pagesize=landscape(A4))
    c.drawImage(img, 5, 2, width=842, height=585)
    c.setFont('Times-Italic', 30, leading=None)
    c.drawCentredString(415, 305, name)
    c.save()
    print('\n\nPDF Generated : '+name)
    sendmail(name, email, file_name)
 

df = pd.read_excel('data.xlsx', sheet_name='papers')
for i in df.index:
    make_certificate(df['Name'][i],df['Email'][i])
print('PDF Generation of Certificates Completed and mail sent.')