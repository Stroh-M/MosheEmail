import imaplib, email, os, re, csv
from email.header import decode_header
from bs4 import BeautifulSoup #type: ignore
from dotenv import load_dotenv #type: ignore

load_dotenv(override=True)

mail = imaplib.IMAP4_SSL("imap.gmail.com")

email_address = os.getenv("EMAIL_ADDRESS")
email_password = os.getenv("EMAIL_PASSWORD")
email_from_1 = os.getenv("EMAIL_FROM_1")
email_from_2 = os.getenv("EMAIL_FROM_2")

mail.login(email_address, email_password)

try:
    print('Login Successful')
    mail.select('INBOX')
    status, result = mail.search(None, f'OR FROM {email_from_1} FROM {email_from_2}')

    email_ids = result[0].split()
    for i in range(len(email_ids)):
        status, msg_data = mail.fetch(email_ids[i], '(RFC822)')
        
        raw_email = email.message_from_bytes(msg_data[0][1])

        subject, encoding = decode_header(raw_email['Subject'])[0]
            
        
        for part in raw_email.walk():
            content_type = part.get_content_type()
        
            if content_type == 'text/html':
                pl = part.get_payload(decode=True)
                soup = BeautifulSoup(pl, 'html.parser')
                
                # Ebay order number scrape 
                for s in soup.find_all('span'):
                    sp_text = s.get_text()
                    e_order_pattern = re.compile(r'^\s*\d{2}-\d{5}-\d{5}\s*$')
                    order = re.findall(pattern=e_order_pattern, string=sp_text)
                    if len(order) > 0:
                        break
                    
                # If no ebay order number look for keurig order number
                # Keurig tracking and order number scrape
                if len(order) <= 0:
                    for x in soup.find_all('td'):
                        td_text = x.get_text()
                        tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
                        order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)
                        tracking = re.findall(pattern=tracking_pattern, string=td_text)
                        order = re.findall(pattern=order_pattern, string=td_text)
                        if len(tracking) > 0 and len(order) > 0:
                            break
                
                # Ebay tracking number scrape
                for y in soup.find_all('p'):
                    p_text = y.get_text().lstrip()
                    e_tracking_pattern = re.compile(r'Tracking number\s*:\s*(\S+)', re.IGNORECASE)
                    tracking = re.findall(pattern=e_tracking_pattern, string=p_text)
                    if len(tracking) > 0:
                        break  
            
        
                # ebay and keurig shipping address scraping 
                shipping_td = soup.find('td', string=lambda t: t and 'Shipping Address' in t)
                e_shipping_h3 = soup.find('h3', string=lambda text: text and 'Your order will ' in text)  
                # Keurig shipping address scrape
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
                    full_address = "\t".join(address_lines)
                elif e_shipping_h3:
                    shipping_p = e_shipping_h3.find_next_sibling('p')
                    if shipping_p:
                        full_address = shipping_p.get_text(separator='\t').strip()




        # with open('C:\\Users\\meir.stroh\\OneDrive\\MosheEmail\\118545073774020185.txt', 'r') as file:
        #     reader = csv.reader(file, delimiter='\t')
        #     rows = list(reader)

        #     for row in rows:
        #         print(row)
        #         if 'Roger Kienast' in row:
        #             print('found')


        print(i)
        print(f'tracking #: {tracking}')
        print(f'order #: {order}')
        print(f'Shipping Address: \n{full_address}')


        
        print('----------END---------')
except imaplib.IMAP4_SSL.error as e:
    print(f'error: {e}')