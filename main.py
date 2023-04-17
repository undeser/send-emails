import smtplib
from email.message import EmailMessage
import pandas as pd

SMTP_SERVER = "smtp.gmail.com" ## gmail
SMTP_PORT = 587
SMTP_USERNAME = "xxx@gmail.com" ## insert email here
SMTP_PASSWORD = "password" ## insert password here

EMAIL_FROM = "xxx@gmail.com" ## insert your email here
EMAIL_SUBJECT = "Email Title" ## email title here

mail = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
mail.starttls()
mail.login(SMTP_USERNAME, SMTP_PASSWORD)

def send_email(EMAIL_TO, name):
    msg = EmailMessage()
    msg['Subject'] = EMAIL_SUBJECT
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
##    msg['Reply-to'] = "xxx@gmail.com" ## optional, for if you want reply to function


### insert body of text here -> can be html

    data = '''
    Dear [insert name],
    '''


    data = data.replace("[insert name]", name) ## replace with name
    msg.set_content(data, subtype='html')
    print(f"EMAIL SENT SUCCESSFULLY TO {name} : {EMAIL_TO}") ## f-string to print each output & ensure email sent
    ##mail.send_message(msg) #### important, to actually send the email

# read pandas csv file
df = pd.read_excel("final.xlsx")

total_rows = len(df)
print(f"TOTAL ROWS: {total_rows}")
counter = 0

for index, row in df.iterrows():
    counter += 1
    # get name get email
    name = row["Name"]
    email = row["Email"]
    # call send email function
    send_email(email, name)
    
mail.quit()
print(f"TOTAL EMAILS SENT: {counter}")
