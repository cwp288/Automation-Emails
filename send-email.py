import ssl
import smtplib
import os
from email.message import EmailMessage
import pandas as pd

# Paths
html_file_path = 'email_template.html'
csv_file_path = 'Startupemails.csv'
resume_path = 'CullenPekResumeEmail.pdf'
#Email Credentials
email_sender = 'pekcullen@gmail.com'
email_password = ''

subject = 'Internship Inquiry: Rutgers University Student'
#Read HTML content
with open(html_file_path, 'r') as file:
    html_content = file.read()
#Read contacts
contacts = pd.read_csv(csv_file_path)

#Create context
context = ssl.create_default_context()

for index, row in contacts.iterrows():
    email_receiver = row['Company Email']
    person_name = row['Person Name']
    company_name = row['Company Name']

    em = EmailMessage()
    em['From'] = email_sender
    em['To'] = email_receiver
    em['Subject'] = subject
    em.set_content(html_content, subtype='html')

    personalized_content = html_content
    personalized_content = personalized_content.replace('{{PERSON_NAME}}', person_name)
    personalized_content = personalized_content.replace('{{COMPANY_NAME}}', company_name)
    personalized_content = personalized_content.replace('{{COMPANY_NAME}}', company_name)
    personalized_content = personalized_content.replace('{{COMPANY_NAME}}', company_name)
    em.set_content(personalized_content, subtype = 'html')

    with open(resume_path, 'rb') as attachment_file:
        file_data = attachment_file.read()
        file_name = os.path.basename(resume_path)

    # Add attachment directly using add_attachment
    em.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com',465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())



print("Emails sent Successfully.")



#  body = """
# Dear _______,
# I hope this email finds you well. My name is Cullen, and I am a current student at Rutgers University, majoring in Computer Science B.S. with a minor in Data Science. 
# I am reaching out to express my interest in internship opportunities within _______, particularly in the field of Software Engineering and Data Science.

# My journey in this field included developing an Automated Job Application Tracker, where I applied object-oriented programming and data structures to manage and organize job applications efficiently. 
# Through this project, I implemented a priority system to categorize applications based on urgency, application date, job title, company, and status. Additionally, I integrated an automation system that scans emails to update an SQL database, with the latest application status. 
# I also enabled the automated removal of denied applications and the generation of organized reports in Excel.

# As a STEM Lead at YMCA Camp Wiggi, I honed my problem-solving skills by organizing and managing diverse activities, ensuring everything ran smoothly and efficiently. 
# I also developed strong conflict resolution abilities, maintaining a positive environment and addressing challenges as they arose. 
# Additionally, I emphasized clean and consistent communication with parents, keeping them informed and addressing any concerns, which reinforced my ability to convey important information effectively.

# I have attached my resume for your reference and would be grateful for the opportunity to discuss how my background and skills align with potential internship roles at ______.

# Thank you for your time and consideration. 

# Best Wishes,

# Cullen Pek.
# """