import emailhandling, os, re, error, datetime, zipfile
from email import utils
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

recipients = [recipient_1, recipient_2]

ebay_tracking_pattern = re.compile(r'Tracking number\s*:\s*(\S+)', re.IGNORECASE)
ebay_order_pattern = re.compile(r'^\s*\d{2}-\d{5}-\d{5}\s*$')
track_order_url_pattern = re.compile(r'\btrack (order|delivery)\b', re.IGNORECASE)

keurig_tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)
keurig_order_pattern = re.compile(r'Order\s*#\s*:\s*(\S+)', re.IGNORECASE)

try:
    mail = emailhandling.Email(email_address, email_password)
    email_ids = mail.get_email_ids('INBOX', email_from_1=email_from_1, email_from_2=email_from_2)
    print('login successful')
    for i in range(len(email_ids)):
        try:
            full_address = None
            order = None
            tracking = None
            
            email_html = mail.get_html(email_id=email_ids[i])

            email_soup = emailhandling.EmailParser(email_html)

            email_date = mail.get_email_date(email_id=email_ids[i])

            tracking = email_soup.find_pattern('p', ebay_tracking_pattern)
            if tracking is None:
                tracking = email_soup.find_pattern('td', keurig_tracking_pattern)

            tracking_href = email_soup.find_pattern('a', track_order_url_pattern)

            order = email_soup.find_pattern('span', ebay_order_pattern)
            if order is None:
                order = email_soup.find_pattern('td', keurig_order_pattern)
            
            full_address = email_soup.get_shipping_address()

            zip = email_soup.get_zip()

            name = email_soup.get_name()

            if tracking is None:
                raise error.No_Tracking_Number(f'<html><p>No tracking number found in email {i}<br /> where customer shipping address is: {full_address}<br /> and the order number is {order}<br /><br /><br />P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
            if order is None:
                raise error.No_Order_Number(f'<html><p>No order number found in email {i}<br />where customer shipping address is: {full_address}<br />and tracking number is {tracking}<br /><br /><br />P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a></html>')
            elif full_address is None or full_address == '':
                raise error.No_Shipping_Address(f'<html><p>No shipping address found in email {i}<br />where order number is: {order}<br />and tracking number is {tracking} as a result can not find the amazon order number<br /><br /><br /> P.S. There might be more issues with this email</p><a href="{tracking_href}">Track Order</a>')
            
            found_match_amazon = False
            found_match = False
            try:
                amazon_orders = emailhandling.File(tsv_file_path, 'tsv')

                reader = amazon_orders.read()
                for row_n, row in enumerate(list(reader)):
                    if len(row) > 0:
                        if name.lower() in row[17].lower().replace('.', '') and zip in row[23]:
                            found_match_amazon = True
                            amazon_order_id = row[0]

                            data = []
                            carrier = emailhandling.get_carrier(tracking_number=tracking[0])
                            ship_date = datetime.datetime.now().date()

                            data.append([(1, amazon_order_id), (4, ship_date), (5, carrier), (7, tracking[0])])

                            try:
                                a_s = emailhandling.File(excel_file_path, 'xlsx', sheet_name)
                                a_s.append_data(data=data)
                                a_s.save()
                                mail.mark_email_as_trash(email_id=email_ids[i])
                            except FileNotFoundError:
                                print(f'Error: No file found at: {excel_file_path}')
                            except PermissionError:
                                print(f'Error: Permission denied most probably cause file open in another program, close file, and try again')
                            except zipfile.BadZipFile:
                                print(f'BadZipFile caught file at {excel_file_path} is not a valid .xlsx (Excel) file')
            except FileNotFoundError:
                print(f'Error: No file found at: {tsv_file_path}')
            except PermissionError:
                print(f'Error: Permission denied to open file at: {tsv_file_path}')
            except OSError as e:
                print(f'An unexpected error occured: {e}')
            if not found_match_amazon:
                try:
                    w_s = emailhandling.File(walmart_order_excel_file, 'xlsx', sheet='Po Details')
                    data = []

                    carrier = emailhandling.get_carrier(tracking_number=tracking[0])

                    data.append([(32, carrier), (33, tracking[0])])
                    for idx, row in enumerate(w_s.iter_rows(), start=1):
                        if zip == row[13] and name == emailhandling.EmailParser.remove_space_from_middle_of_string(None, row[5]):
                            found_match = True
                            w_s.fill_data(idx, data=data)
                    w_s.save()
                    mail.mark_email_as_trash(email_id=email_ids[i])
                except FileNotFoundError:
                    print(f'Error: No file found at: {walmart_order_excel_file}')
                except PermissionError:
                    print(f'Error: Permission denied most probably cause file open in another program, close file, and try again')
                except zipfile.BadZipFile:
                    print(f'BadZipFile caught file at {walmart_order_excel_file} is not a valid .xlsx (Excel) file')
                except Exception as e:
                    print(f'Unexpected error: {e}')

            if not found_match:
                error_message = 'Did not find match in walmart nor amazon'
                try:
                    e_s = emailhandling.File(error_excel_path, 'xlsx')

                    error_data = []

                    error_data.append([(1, error_message), (2, order[0]), (3, tracking[0]), (4, full_address), (5, name), (6, zip), (7, email_date), (8, datetime.datetime.now())])
                    e_s.append_data(data=error_data)
                    e_s.save()
                except FileNotFoundError:
                    print(f'Error: No file found at: {error_excel_path}')
                except PermissionError:
                    print(f'Error: Permission denied most probably cause file open in another program, close file, and try again')
                except zipfile.BadZipFile:
                    print(f'BadZipFile caught file at {error_excel_path} is not a valid .xlsx (Excel) file')
                except Exception as e:
                    print(f'Unexpected error: {e}')
        except error.No_Order_Number as non_e:
            mail.send_message('No Order Number', recipients=recipients, email_msg=str(non_e))
        except error.No_Shipping_Address as nsa_e:
            mail.send_message('No Shipping Address', recipients, str(nsa_e))
        except error.No_Tracking_Number as ntn_e:
            mail.send_message('No Tracking Number', recipients, str(ntn_e))
        print(f'---------- Processed email #{i} ---------')
    mail.close_mails()
    w_s.convert_file_type(shipping_txt_file)
    print('Logged out successfully')
except emailhandling.imaplib.IMAP4_SSL.error as mail_s_e:
    print(f'error: {mail_s_e}')
except Exception as e:
    print(f'Unexpected error: {e}')