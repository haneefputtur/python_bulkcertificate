# python_bulkcertificate
Creating Bulk Certificates for event using python

More info : https://haneefputtur.com/creating-event-certificates-in-python.html

Features of this script:

Uses Excel sheet as database of users – Only Name and Email column is required.

Can be used to generate pdf files without sending email.

Gmail can be used to send out mails.

Any other smtp mails can be configured to send out.

Pre Requisites:

Data in Excel sheet as per the sample

Enable 2 step authentication in google and create separate app password to use with python. Or else smtp will not work.

Certificate background template with all logo and signatures should be prepared in photoshop and saved as jpg.

Python 3.6+ installed

Following python packages must be installed

# pandas

#smtplib

#reportlab [pip install reportlab]
