import imaplib
import email
import requests
from datetime import datetime
import urllib3
import re
import chardet
import pytz
from dateutil.parser import parse
import json
import os

mail_host = 'imap.exmail.qq.com'
mail_port = 993
mail_username = 'xxx@example.com'
mail_password = 'xxxxxx'

mail = imaplib.IMAP4_SSL(mail_host, mail_port)

mail.login(mail_username, mail_password)
mail.select('INBOX')

current_date = datetime.now(pytz.timezone('Asia/Shanghai')).strftime("%d-%b-%Y")

search_criteria = f'ON "{current_date}" FROM "no-reply@loudongyun.360.cn"'.encode('utf-8')
result, data = mail.search(None, search_criteria)
mail_ids = data[0].split()

print(f'Found {len(mail_ids)} mails') 

saved_emails_file = 'saved_emails.json'
if os.path.exists(saved_emails_file):
    with open(saved_emails_file, 'r') as f:
        saved_emails = json.load(f)
else:
    saved_emails = []

new_emails = []
if mail_ids:
    for mail_id in mail_ids:
        _, data = mail.fetch(mail_id, '(RFC822)')
        raw_email = data[0][1]
        email_message = email.message_from_bytes(raw_email)
        from_address = str(email_message.get('From')) 

        if 'no-reply@loudongyun.360.cn' not in from_address:
            continue
        
        email_date = email_message.get('Date')
        email_date = parse(email_date).astimezone(pytz.timezone('Asia/Shanghai'))
        if email_date.strftime("%d-%b-%Y") != current_date:
            continue

        title = email_message.get('Subject')
        content = None
        if email_message.is_multipart():
            for part in email_message.get_payload():
                if part.get_content_type() == 'text/plain':
                    content = part.get_payload(decode=True)
        else:
            content = email_message.get_payload(decode=True)

        if content is not None:
            encoding = chardet.detect(content)['encoding']
            content = content.decode(encoding)
            print(type(content))
            url_regex = r'(http://loudongyun\.360\.net/bug/detail/[^\s]+)'
            url_match = re.search(url_regex, content)
            if url_match:
                url = url_match.group(0).split('">')[0] 
            else:
                url = ""
            level_regex = r'漏洞等级：(\w+)'
            level_match = re.search(level_regex, content)
            if level_match:
                level = level_match.group(1)  
            else:
                level = ""

            if title is not None:
                decoded_header = email.header.decode_header(title)[0]
                if decoded_header[1] is not None:
                    title = decoded_header[0].decode(decoded_header[1])
                else:
                    title = decoded_header[0]

            if {'title': title, 'url': url, 'level': level} not in saved_emails:
                new_emails.append({'title': title, 'url': url, 'level': level})

if new_emails:
    for email in new_emails:
        webhook_url = 'xxxxx'
        headers = {'Content-Type': 'application/json'}
        payload = {
            'msgtype': 'text',
            'text': {
                'content': f'标题：{email["title"]}\n漏洞等级：{email["level"]}\n链接：{email["url"]}'
            }
        }

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.post(webhook_url, json=payload, headers=headers)
        print('推送结果:', response.text)

    saved_emails.extend(new_emails)
    with open(saved_emails_file, 'w') as f:
        json.dump(saved_emails, f)

mail.logout()
