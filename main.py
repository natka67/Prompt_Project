import requests
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

sender_email = ''
sender_password = ''
recipient_email = 'nacia354@gmail.com'
subject = '[Decision]  Request for Your Input on Report Usage Discrepancies'

data = pd.read_excel('data.xlsx')
categories = dict({
    1: 'Zero users of the report',
    2: 'Quarterly usage of the report was less than declared',
    3: 'Monthly usage of the report was less than declared'
})

owners = list(set(data['Business Owner']))
mails = {}

for owner in owners:
    reports = data[data['Business Owner'] == owner]
    part1 = f'''Dear {owner},

    I trust this message finds you well. 
    
    Upon completing our MMC, we've detected disparities in specific reports where the actual usage is falling short 
    of the declared number. I wanted to bring these findings to your attention:\n'''

    part2a = f'\tZero users of the report: {', '.join(set(reports[reports['Kategoria'] == 1]['Report Name']))}\n' if len(set(reports[reports['Kategoria'] == 1]['Report Name']))!=0 else ''
    part2b = f'\tQuarterly usage less than declared: {', '.join(set(reports[reports['Kategoria'] == 2]['Report Name']))}\n' if len(set(reports[reports['Kategoria'] == 2]['Report Name']))!=0 else ''
    reports2c = reports[reports['Kategoria'] == 3]
    reports2c = reports2c.groupby('Report Name')['Month'].agg(lambda x: ', '.join(x)).reset_index()
    reports2c['Report_Month'] = reports2c['Report Name'] + ' (' + reports2c['Month'] + ')'
    part2c = f'\tMonthly usage less than declared: {', '.join(set(reports2c['Report_Month']))}\n' if len(set(reports2c['Report_Month'])) != 0 else ''
    part3 = f'''\nTo effectively address this issue, we seek your guidance on whether these reports should undergo decommissioning, upgrading, or if adjustments to the expected number of views per month are necessary.

    Here is a detailed breakdown of the expected month views for your reference:
    {reports[['Report Name', 'Expected Month Views Number']].drop_duplicates(subset=['Report Name', 'Expected Month Views Number'])}
    \nPlease review the information and provide your input at your earliest convenience.
    
    Kind regards,
    Natalia Adamczyk'''

    message_body = ''.join([part1, part2a, part2b, part2c, part3])

    mails[owner] = message_body
    print(reports.to_json())
    print(message_body)
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the message body
    msg.attach(MIMEText(message_body, 'plain'))

    # Connect to the SMTP server
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            # Log in to your Outlook account
            server.login(sender_email, sender_password)

            # Send the email
            server.sendmail(sender_email, recipient_email, msg.as_string())

            print('Email sent successfully.')
    except Exception as e:
        print(f'Error: {e}')