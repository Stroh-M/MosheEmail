import error, email_handling, re, os, email_utils_beta, logging, zipfile, datetime, traceback
from dotenv import load_dotenv #type: ignore

load_dotenv(override=True)

email_address = os.getenv("EMAIL_ADDRESS")
email_password = os.getenv("EMAIL_PASSWORD")
email_from_1 = os.getenv("EMAIL_FROM_1")
email_from_2 = os.getenv("EMAIL_FROM_2")
recipient_1 = os.getenv("RECIPIENT_1")
recipient_2 = os.getenv("RECIPIENT_2")
excel_file_path = os.getenv("EXCEL_FILE_PATH")
tsv_file_path = os.getenv("TSV_FILE_PATH")
sheet_name = os.getenv("SHEET_NAME")
error_excel_path = os.getenv("ERROR_EXCEL_PATH")
shipping_txt_file = os.getenv("SHIPPING_TXT_FILE")
walmart_order_excel_file = os.getenv("WALMART_ORDER_EXCEL_FILE")


mail = email_handling.Email(email_address, email_password)
email_ids = mail.get_email_ids('INBOX', email_from_1=email_from_1, email_from_2=email_from_2)

def write_to_excel(path, sheet_name, data):
    try:
        a_s = email_handling.File(path=path, type='xlsx', sheet=sheet_name)
        a_s.append_data(data=data)
        a_s.save()
    except FileNotFoundError:
        print(f'Error: No file found at: {path}')
    except PermissionError:
        print(f'Error: Permission denied most probably cause file open in another program, close file, and try again')
    except zipfile.BadZipFile:
        print(f'BadZipFile caught file at {path} is not a valid .xlsx (Excel) file')
        
def handle_amazon_orders(txt_path, name, zip, tracking, excel_path, sheet):
    a_o = email_handling.File(path=txt_path, type='tsv')
    name_index = a_o.find_column_index('recipient-name')
    zip_index = a_o.find_column_index('ship-postal-code')
    order_id_idx = a_o.find_column_index('order-id')
    try:
        for row_n, row in enumerate(a_o.read()):
            if len(row) > 0:
                if name.lower().replace('.', '') in email_handling.EmailParser.remove_space_from_middle_of_string(None, string=row[name_index]) and zip in row[zip_index]:
                    global found_match_amazon
                    found_match_amazon = False
                    amazon_order_id = row[order_id_idx]
                    
                    data = []
                    
                    carrier = email_utils_beta.get_carrier(tracking_number=tracking)
                    ship_date = datetime.datetime.now().date()
                    
                    a_s = email_handling.File(path=excel_path, type='xlsx', sheet='ShippingConfirmation')
                    
                    a_e_order_id_idx = a_s.find_column_index('order-id')
                    a_e_ship_date_idx = a_s.find_column_index('ship-date')
                    a_e_carrier_code_idx = a_s.find_column_index('carrier-code')
                    a_e_tracking_number_idx = a_s.find_column_index('tracking-number')
                    data.append([(a_e_order_id_idx, amazon_order_id), (a_e_ship_date_idx, ship_date), (a_e_carrier_code_idx, carrier), (a_e_tracking_number_idx, tracking)])
                    
                    write_to_excel(path=excel_path, sheet_name=sheet, data=data)   
    except FileNotFoundError:
        print(f'Error: No file found at: {txt_path}')
    except PermissionError:
        print(f'Error: Permission denied to open file at: {txt_path}')
    except OSError as e:
        print(f'An unexpected error occured: {e}')
        print(f'Traceback: {traceback.format_exc()}')
        
def handle_walmart_orders(path, name, zip, tracking, sheet):
    try:
        w_s = email_handling.File(path=path, type='xlsx', sheet='Po Details')
        data = []
        
        walmart_name_idx = w_s.find_column_index('Customer Name')
        walmart_zip_idx = w_s.find_column_index('Zip')
        walmart_update_tracking_idx = w_s.find_column_index('Update Tracking Number')
        walmart_update_carrier = w_s.find_column_index('Update Carrier')
        
    except:
        pass 
        
        
    

def proccess_email(mail, email_ids, id):
    ebay_tracking_pattern = re.compile(r'Tracking number\s*:\s*(\S+)', re.IGNORECASE)
    ebay_order_pattern = re.compile(r'^\s*\d{2}-\d{5}-\d{5}\s*$')
    track_order_url_pattern = re.compile(r'\btrack (order|delivery)\b', re.IGNORECASE)
    track_order_url_backup = re.compile(r'\btrack my order\b', re.IGNORECASE)
    keurig_tracking_pattern_a = re.compile(r'Tracking Number:\s*([A-Za-z0-9]+)')
    keurig_tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
    keurig_order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)

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
        
        handle_amazon_orders(txt_path=tsv_file_path, name=name, zip=zip, tracking=tracking[0], excel_path=excel_file_path, sheet=sheet_name)
    except:
        pass
    
for i in range(len(email_ids)):    
    proccess_email(mail=mail, email_ids=email_ids, id=i)