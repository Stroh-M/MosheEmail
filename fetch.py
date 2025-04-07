import imaplib, email, os, re
from email.header import decode_header
from bs4 import BeautifulSoup #type: ignore
from dotenv import load_dotenv #type: ignore

load_dotenv(override=True)

mail = imaplib.IMAP4_SSL("imap.gmail.com")

email_address = os.getenv("EMAIL_ADDRESS")
email_password = os.getenv("EMAIL_PASSWORD")
email_from = os.getenv("EMAIL_FROM")

mail.login(email_address, email_password)

try:
    print('Login Successful')
    mail.select('INBOX')
    status, result = mail.search(None, 'From', f'{email_from}')

    email_ids = result[0].split()
    for i in range(len(email_ids)):
        status, msg_data = mail.fetch(email_ids[i], '(RFC822)')
        
        raw_email = email.message_from_bytes(msg_data[0][1])

        subject, encoding = decode_header(raw_email['Subject'])[0]

        print(subject)

        for part in raw_email.walk():
            content_type = part.get_content_type()

        
            # Keurig email scraping
            if content_type == 'text/html':
                pl = part.get_payload(decode=True)
                soup = BeautifulSoup(pl, 'html.parser')
                
                # print(pl)
                for x in soup.find_all('td'):
                    td_text = x.get_text()
                    tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
                    order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)
                    tracking = re.findall(pattern=tracking_pattern, string=td_text)
                    order = re.findall(pattern=order_pattern, string=td_text)
                    if len(tracking) > 0 and len(order) > 0:
                        break

        shipping_td = soup.find("td", string=lambda t: t and 'Shipping Address' in t)  

        if shipping_td:
            parent_tr = shipping_td.find_parent('tr')
            address_table = parent_tr.find_parent('table')
            all_rows = address_table.find_all('tr')

            start_index = all_rows.index(parent_tr) + 1
            address_lines = []

            for row in all_rows[start_index:]:
                tds = row.find_all('td')
                for td in tds:
                    text = td.get_text()
                    text = text.lstrip()
                    if text not in address_lines:
                        address_lines.append(text)

            address_lines = [item for item in address_lines if item != '']
            full_address = "\n".join(address_lines)

        
            print(f'tracking #: {tracking[0]}')
            print(f'order #: {order[0]}')
            print(f'Shipping Address: \n{full_address}')


        
        print('----------END---------')
except imaplib.IMAP4_SSL.error as e:
    print(f'error: {e}')