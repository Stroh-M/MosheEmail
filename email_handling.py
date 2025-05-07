import email.utils
import pandas as pd #type: ignore
import imaplib, smtplib, email, pytz, re, openpyxl, csv, traceback, zipfile, logging, inspect #type: ignore
from bs4 import BeautifulSoup #type: ignore
from email.message import EmailMessage

logger = logging.getLogger(__name__)

class Email():
    def __init__(self, emailAdress, emailPassword):
        try:
            self.email_address = emailAdress
            self.imap_mail = imaplib.IMAP4_SSL(f'imap.gmail.com')
            self.smtp_mail = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            self.imap_mail.login(emailAdress, emailPassword)
            self.smtp_mail.login(emailAdress, emailPassword)
        except Exception:
            logger.exception(f'Failed to initialize {self.__class__.__name__}')
            
    def get_email_ids(self, mailbox, email_from_1, email_from_2):
        try:
            self.imap_mail.select(mailbox=mailbox)
            status, self.result = self.imap_mail.search(None, f'OR FROM {email_from_1} FROM {email_from_2}')
            return  self.result[0].split()
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
            
    def __convert_to_message(self, email_id, format='RFC822'):
        try:
            status, msg_data = self.imap_mail.fetch(email_id, f'{format}')
            raw_email = email.message_from_bytes(msg_data[0][1])
            return raw_email
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def get_html(self, email_id, format='RFC822'):
        try:
            raw_email = self.__convert_to_message(email_id, f'{format}')
            for part in raw_email.walk():
                content_type = part.get_content_type()
                if content_type == 'text/html':
                    pl = part.get_payload(decode=True)
            return pl
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def get_email_date(self, email_id, format='RFC822', local_tz='America/New_York'):
        try:
            local_tz = pytz.timezone(local_tz)
            raw_email = self.__convert_to_message(email_id=email_id, format=format)
            date_string = raw_email.get('Date')
            email_date = email.utils.parsedate_to_datetime(date_string).astimezone(local_tz).replace(tzinfo=None)
            return email_date
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def __remove_html_tags(self, message):
        try:
            converted_breaks = re.sub(r'<br\s*/?>', '\n', message)
            cleaned_string = re.sub(r'<[^>]+>', '', converted_breaks)
            return cleaned_string
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def send_message(self, subject, recipients, email_msg):
        try:
            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = self.email_address
            msg['To'] = ', '.join(recipients)
            msg.set_content(self.__remove_html_tags(email_msg))
            msg.add_alternative(email_msg, subtype='html')
            self.smtp_mail.send_message(msg=msg)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def mark_email_as_trash(self, email_id):
        try:
            self.imap_mail.store(email_id, '+X-GM-LABELS', '\\Trash')
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def close_mails(self):
        try:
            if self.imap_mail:
                self.imap_mail.select('INBOX')
                self.imap_mail.close()
                self.imap_mail.logout()
            if self.smtp_mail:
                self.smtp_mail.quit()
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

class EmailParser():
    def __init__(self, email_html, parser='html.parser'):
        try:
            self.soup = BeautifulSoup(email_html, parser)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
            
    def prettify(self):
        try:
            return self.soup.prettify()
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    # If a is passed as element it will get the href for that a element 
    def find_pattern(self, element, pattern, href=False):
        try:
            if href:
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
                    if found_e:
                        return found_e
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
                
    def __k_shipping(self):
        try:
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
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def __k_shipping_h4(self):
        try:
            shipping_h4 = self.soup.find('h4', string=lambda t: t and 'Shipping Address' in t)
            shipping_p = shipping_h4.find_next_sibling('p')
            if shipping_p:
                full_address = shipping_p.get_text(separator='\t').strip()
                return full_address
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
        
    def __e_shipping(self):
        try:
            shipping_h3 = self.soup.find('h3', string=lambda t: t and 'Your order will ' in t)
            shipping_p = shipping_h3.find_next_sibling('p')
            if shipping_p:
                full_address = shipping_p.get_text(separator='\t').strip()
                return full_address
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def get_shipping_address(self):
        try:
            if self.soup.find('td', string=lambda t: t and 'Shipping Address' in t):
                self.shipping_address = self.__k_shipping()
            elif self.soup.find('h3', string=lambda t: t and 'Your order will ' in t):
                self.shipping_address = self.__e_shipping()
            elif self.soup.find('h4', string=lambda t: t and 'Shipping Address' in t):
                self.shipping_address = self.__k_shipping_h4()
            else:
                self.shipping_address = None
            return self.shipping_address
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def get_back_up_tracking(self):
        try:
            tracking = []
            found_ = self.find_element('span', 'Number')

            parent_dt = found_.find_parent()
            parent_div = parent_dt.find_parent()

            spans = parent_div.find_all('span')
            tracking.append(spans[-1].get_text())
            return tracking   
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')     
        
    def remove_space_from_middle_of_string(self, string):
        try:
            clean_string = re.sub(r'\s+', ' ', string=string)
            return clean_string
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
        
    def get_name(self):
        try:
            address = self.get_shipping_address()
            if address is not None:
                self.address = re.split(r'\t+', self.shipping_address)
                name = self.remove_space_from_middle_of_string(self.address[0])
            else: 
                name = ''
            return name
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def find_element(self, element, string):
        try:
            found_element = self.soup.find(element, string=lambda t: t and string in t)
            return found_element
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def get_zip(self):
        try:
            address = self.get_shipping_address()
            if address is not None:
                zip_code_pattern = re.compile(r'\b(\d{5})(?:-\d{4})?\b')
                zip_code = re.findall(pattern=zip_code_pattern, string=self.shipping_address)
                zip = zip_code[-1]
            else:
                zip = ''
            return zip
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    

class File():
    def __init__(self, path, type, sheet='Sheet1', mode='r'):
        try:
            self.type = type
            self.file_path = path
            self.sheet_name = sheet
            if type == 'xlsx':
                self.workbook = openpyxl.load_workbook(self.file_path)
                self.sheet = self.workbook[sheet]
            elif type in ('txt', 'tsv'):
                self.doc = open(self.file_path, mode=mode)
        except Exception:
            logger.exception(f'Failed to initialize {self.__class__.__name__} with file: {path}')

    def read(self, delimiter='\t'):
        try:
            return csv.reader(self.doc, delimiter=delimiter)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
    
    def get_max_row(self):
        try:
            return self.sheet.max_row
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def iter_rows(self, values_only=True):
        try:
            return self.sheet.iter_rows(values_only=values_only)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def append_data(self, data):
        try:
            max_row = self.get_max_row()
            for idx, row_data in enumerate(data, start=max_row+1):
                for col_idx, value in row_data:
                    self.sheet.cell(row=idx, column=col_idx, value=value)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
                
    def fill_data(self, row_num, data):
        try:
            for row in data:
                for col, value in row:
                    self.sheet.cell(row=row_num, column=col, value=value)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def save(self):
        try:
            self.workbook.save(self.file_path)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')

    def convert_file_type(self, new_file_path, to_type='tsv'):
        try:
            if self.type == 'xlsx' and to_type in ('tsv', 'csv'):
                file_to_convert = pd.read_excel(self.file_path, engine='openpyxl', sheet_name=self.sheet_name)
                file_to_convert.to_csv(new_file_path, sep=f'{',' if to_type == 'csv' else ('\t' if to_type == 'tsv' else ',')}')
            else:
                return NotImplementedError('Can only convert .xlsx to either .tsv or .csv for now')
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
            
    def find_column_index(self, column_title):
        try:
            if self.type == 'xlsx':
                for idx, col in enumerate(self.sheet.iter_cols(max_row=1, values_only=True), start=1):
                    if column_title == col[0]:
                        return idx
            elif self.type in ('tsv', 'csv'):
                reader = self.read()
                for idx, row in enumerate(reader):
                    if idx == 1:
                        break
                    for c_idx, col in enumerate(row):
                        if col.strip() == column_title:
                            return c_idx
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')
        
    def delete_all_cells(self):
        try:
            max_row = self.get_max_row()
            self.sheet.delete_rows(idx=2, amount=max_row)
        except Exception:
            logger.exception(f'Error in: {inspect.currentframe().f_code.co_name}')