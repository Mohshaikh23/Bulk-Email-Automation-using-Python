import pandas as pd
import smtplib as sm
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

df = pd.read_csv("email trial.csv")
emails = list(df['email'])
message = list(df['Message'])
try:
    server = sm.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login("mohsin.shaikh324@gmail.com","Mohsin@007")
    from_ = "mohsin.shaikh324@gmail.com"
    to_= emails
    message = MIMEText()
    message['Subject'] = "testing Automation"
    message['from'] = "mohsin.shaikh324@gmail.com"


except Exception as e:
    print(e)