import email.utils
import imaplib, smtplib, email, pytz, re, openpyxl, csv #type: ignore
from bs4 import BeautifulSoup #type: ignore


class Email():
    def __init__(self, emailAdress, emailPassword):
        self.imap_mail = imaplib.IMAP4_SSL(f'imap.gmail.com')
        self.smtp_mail = smtplib.SMTP_SSL(f'smtp.gmail.com', 465)
        self.imap_mail.login(emailAdress, emailPassword)
        self.smtp_mail.login(emailAdress, emailPassword)

    def get_email_ids(self, mailbox, email_from_1, email_from_2):
        self.imap_mail.select(mailbox=mailbox)
        status, self.result = self.imap_mail.search(None, f'OR FROM {email_from_1} FROM {email_from_2}')
        return  self.result[0].split()
        
    def __convert_to_message(self, email_id, format='RFC822'):
        status, msg_data = self.imap_mail.fetch(email_id, f'{format}')
        raw_email = email.message_from_bytes(msg_data[0][1])
        return raw_email
    
    def get_html(self, email_id, format='RFC822'):
        raw_email = self.__convert_to_message(email_id, f'{format}')
        for part in raw_email.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html':
                pl = part.get_payload(decode=True)
        return pl
    
    def get_email_date(self, email_id, format='RFC822', local_tz='America/New_York'):
        local_tz = pytz.timezone(local_tz)
        raw_email = self.__convert_to_message(email_id=email_id, format=format)
        date_string = raw_email.get('Date')
        email_date = email.utils.parsedate_to_datetime(date_string).astimezone(local_tz).replace(tzinfo=None)
        return email_date
    
    def close_mails(self):
        self.imap_mail.close()
        self.imap_mail.logout()
        self.smtp_mail.quit()

class EmailParser():
    def __init__(self, email, parser='html.parser'):
        self.soup = BeautifulSoup(email, parser)

    def prettify(self):
        return self.soup.prettify()
    
    # If a is passed as element it will get the href for that a element 
    def find_pattern(self, element, pattern):
        if element == 'a':
            for a in self.soup.find_all(element):
                a_text = a.get_text()
                found_a = re.findall(pattern=pattern, string=a_text)
                if found_a:
                    href = a.get('href')
                    return href
        else:
            for e in self.soup.find_all(element):
                e_text = e.get_text()
                found_e = re.findall(pattern=pattern, string=e_text)
                if len(found_e) > 0:
                    return found_e
                
    def __k_shipping(self):
        shipping_td = self.soup.find('td', string=lambda t: t and 'Shipping Address' in t)
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
        return full_address
    
    def __e_shipping(self):
        shipping_h3 = self.soup.find('h3', string=lambda t: t and 'Your order will ' in t)
        shipping_p = shipping_h3.find_next_sibling('p')
        if shipping_p:
            full_address = shipping_p.get_text(separator='\t').strip()
            return full_address

    def get_shipping_address(self):
        if self.soup.find('td', string=lambda t: t and 'Shipping Address' in t):
            self.shipping_address = self.__k_shipping()
            return self.shipping_address
        elif self.soup.find('h3', string=lambda t: t and 'Your order will ' in t):
            self.shipping_address = self.__e_shipping()
            return self.shipping_address
        else:
            return None
        
    def get_name(self):
        self.get_shipping_address()
        self.address = re.split(r'\t+', self.shipping_address)
        name = re.sub(r'\s+', ' ', self.address[0]).strip()
        return name
    
    def get_zip(self):
        self.get_shipping_address()
        zip_code_pattern = re.compile(r'\b(\d{5})(?:-\d{4})?\b')
        zip_code = re.findall(pattern=zip_code_pattern, string=self.shipping_address)
        zip = zip_code[-1]
        return zip
    

class File():
    def __init__(self, path, type, sheet='Sheet1', mode='r'):
        self.file_path = path
        if type == 'xlsx':
            self.workbook = openpyxl.load_workbook(self.file_path)
            self.sheet = self.workbook[sheet]
        elif type in ('txt', 'tsv'):
            self.doc = open(self.file_path, mode=mode)


    def read(self):
        return csv.reader(self.doc, delimiter='\t')
    
    def get_max_row(self):
        return self.sheet.max_row


    def append_data(self, data):
        max_row = self.get_max_row()
        for idx, row_data in enumerate(data, start=max_row+1):
            print(row_data)
            for col_idx, value in row_data:
                print(idx)
                self.sheet.cell(row=idx, column=col_idx, value=value)
        

    def save(self):
        self.workbook.save(self.file_path)
            
        




