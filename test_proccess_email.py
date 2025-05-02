import error, email_handling, re, os
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


def proccess_email(mail, email_ids, id):
    ebay_tracking_pattern = re.compile(r'Tracking number\s*:\s*(\S+)', re.IGNORECASE)
    ebay_order_pattern = re.compile(r'^\s*\d{2}-\d{5}-\d{5}\s*$')
    track_order_url_pattern = re.compile(r'\btrack (order|delivery)\b', re.IGNORECASE)

    keurig_tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
    keurig_order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)

    try:
        email_html = mail.get_html(email_id=email_ids[id])

        email_soup = email_handling.EmailParser(email_html)

        email_date = mail.get_email_date(email_id=email_ids[id])

        tracking = email_soup.find_pattern('p', ebay_tracking_pattern)
        if tracking is None:
            tracking = email_soup.find_pattern('td', keurig_tracking_pattern)

        tracking_href = email_soup.find_pattern('a', track_order_url_pattern)

        order = email_soup.find_pattern('span', ebay_order_pattern)
        if order is None:
            order = email_soup.find_pattern('td', keurig_order_pattern)
                
        full_address = email_soup.get_shipping_address()
        print(full_address)

        zip = email_soup.get_zip()
        print(zip)

        name = email_soup.get_name()

        if tracking is None:
            tracking = email_handling.get_backup_tracking(tracking_href)
            if tracking is None:
                raise error.No_Tracking_Number(f'<html><p>No tracking number found in email {id}<br /> where customer shipping address is: {full_address}<br /> and the order number is {order}<br /><br /><br />P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
        elif order is None:
            raise error.No_Order_Number(f'<html><p>No order number found in email {id}<br />where customer shipping address is: {full_address}<br />and tracking number is {tracking}<br /><br /><br />P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a></html>')
        elif full_address is None or full_address == '':
            raise error.No_Shipping_Address(f'<html><p>No shipping address found in email {id}<br />where order number is: {order}<br />and tracking number is {tracking} as a result can not find the amazon order number<br /><br /><br /> P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
        
        return tracking_href, email_date, name, zip, order[0], tracking[0]
    except:
        pass
    
for i in range(len(email_ids)):    
    values_ = proccess_email(mail=mail, email_ids=email_ids, id=i)
    print(values_, flush=True)