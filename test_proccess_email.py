import error, email_handling, re, os, email_utils_beta, logging, datetime
from dotenv import load_dotenv #type: ignore

load_dotenv(override=True)

recipient_1 = os.getenv("RECIPIENT_1")
recipient_2 = os.getenv("RECIPIENT_2")
excel_file_path = os.getenv("EXCEL_FILE_PATH")
tsv_file_path = os.getenv("TSV_FILE_PATH")
sheet_name = os.getenv("SHEET_NAME")
error_excel_path = os.getenv("ERROR_EXCEL_PATH")
shipping_txt_file = os.getenv("SHIPPING_TXT_FILE")
walmart_order_excel_file = os.getenv("WALMART_ORDER_EXCEL_FILE")
        
def handle_amazon_orders(txt_path, name, zip, tracking, excel_path, sheet):
    try:
        found_match_amazon = False
        a_o = email_handling.File(txt_path, type='tsv')
        print(name)
        for row_n, row in enumerate(a_o.read()):
            if len(row) > 0:
                if name.lower().replace('.', '') in email_handling.EmailParser.remove_space_from_middle_of_string(None, string=row[17].lower().replace('.', '')) and zip in row[23]:
                    
                    found_match_amazon = True
                    amazon_order_id = row[0]
                    
                    data = []
                    
                    carrier = email_utils_beta.get_carrier(tracking_number=tracking)
                    ship_date = datetime.datetime.now().date()
                    
                    data.append([(1, amazon_order_id), (4, ship_date), (5, carrier), (7, tracking)])

                    a_s = email_handling.File(path=excel_path, type='xlsx', sheet=sheet)
                    
                    a_s.append_data(data=data)
                    a_s.save()
                    
        if found_match_amazon:
            return True
        else:
            return False
    except Exception as e:
        print(f'Error: {e}')
        
def handle_walmart_orders(path, name, zip, tracking, sheet):
    try:
        found_match = False
        w_s = email_handling.File(path=path, type='xlsx', sheet=sheet)
        data = []
        
        carrier = email_utils_beta.get_carrier(tracking_number=tracking)
        
        data.append([(37, carrier), (38, tracking)])
        
        for idx, row in enumerate(w_s.iter_rows(), start=1):
            if zip == row[13] and name == email_handling.EmailParser.remove_space_from_middle_of_string(None, string=row[5]):
                found_match = True
                w_s.fill_data(idx, data=data)
        w_s.save()
        
        if found_match:
            return True
        else:
            return False
    except Exception as e:
        print(f'Error: {e}') 
        
def proccess_email(mail, email_ids, id):
    ebay_tracking_pattern = re.compile(r'Tracking number\s*:\s*(\S+)', re.IGNORECASE)
    ebay_order_pattern = re.compile(r'^\s*\d{2}-\d{5}-\d{5}\s*$')
    track_order_url_pattern = re.compile(r'\btrack (order|delivery)\b', re.IGNORECASE)
    track_order_url_backup = re.compile(r'\btrack my order\b', re.IGNORECASE)
    keurig_tracking_pattern_a = re.compile(r'Tracking Number:\s*([A-Za-z0-9]+)')
    keurig_tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
    keurig_order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)
    
    recipients = [recipient_1, recipient_2]

    try:
        email_html = mail.get_html(email_id=email_ids[id])

        email_soup = email_handling.EmailParser(email_html)

        email_date = mail.get_email_date(email_id=email_ids[id])

        tracking = email_soup.find_pattern('p', ebay_tracking_pattern)
        if tracking is None:
            tracking = email_soup.find_pattern('td', keurig_tracking_pattern)
            
        if tracking is None:
            tracking = email_soup.find_pattern('td', pattern=keurig_tracking_pattern_a)
          
        tracking_href = email_soup.find_pattern('a', track_order_url_pattern, href=True)
        if tracking_href is None:
            tracking_href = email_soup.find_pattern('a', pattern=track_order_url_backup, href=True)

        order = email_soup.find_pattern('span', ebay_order_pattern)
        if order is None:
            order = email_soup.find_pattern('td', keurig_order_pattern)
            
        full_address = email_soup.get_shipping_address()
        print(full_address)

        zip = email_soup.get_zip()
        print(zip)

        name = email_soup.get_name()

        if tracking is None:
            tracking = email_utils_beta.get_backup_tracking(tracking_href)
            if tracking is None:
                raise error.No_Tracking_Number(f'<html><p>No tracking number found in email {id}<br /> where customer shipping address is: {full_address}<br /> and the order number is {order}<br /><br /><br />P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
        elif full_address is None or full_address == '':
            raise error.No_Shipping_Address(f'<html><p>No shipping address found in email {id}<br />where order number is: {order}<br />and tracking number is {tracking} as a result can not find the amazon order number<br /><br /><br /> P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
        
        result_amazon = handle_amazon_orders(txt_path=tsv_file_path, name=name, zip=zip, tracking=tracking[0], excel_path=excel_file_path, sheet=sheet_name)
        if not result_amazon:
            result_walmart = handle_walmart_orders(walmart_order_excel_file, name=name, zip=zip, tracking=tracking[0], sheet='Po Details')
            if not result_walmart:
                print(f'{name}, {zip}, {order[0]}, couldn''t find match in Amazon or Walmart')
    except error.No_Shipping_Address as nsa_e:
        mail.send_message('No Shipping Address', recipients, str(nsa_e))
    except error.No_Tracking_Number as ntn_e:
        mail.send_message('No Tracking Number', recipients, str(ntn_e))   
    except Exception as e:
        print(f'Error: {e}')
    
def clear_amazon_shippment_excel(path, sheet):
    try:
        a_ws = email_handling.File(path=path, type='xlsx', sheet=sheet)
        a_ws.delete_all_cells()
        a_ws.save()
    except Exception as e:
        print(f'Error: {e}')
    
def main():
    email_address = os.getenv("EMAIL_ADDRESS")
    email_password = os.getenv("EMAIL_PASSWORD")
    email_from_1 = os.getenv("EMAIL_FROM_1")
    email_from_2 = os.getenv("EMAIL_FROM_2")
    excel_file_path = os.getenv("EXCEL_FILE_PATH")
    sheet_name = os.getenv("SHEET_NAME")
        
    mail = email_handling.Email(email_address, email_password)
    email_ids = mail.get_email_ids('INBOX', email_from_1=email_from_1, email_from_2=email_from_2)
        
    clear_amazon_shippment_excel(excel_file_path, sheet=sheet_name)
    for i in range(len(email_ids)):    
        proccess_email(mail=mail, email_ids=email_ids, id=i)
        print(f'---------process email {i}')
            
main()