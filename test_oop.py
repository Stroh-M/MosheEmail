import oop, os, re
from dotenv import load_dotenv #type:ignore

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

track_order_link_pattern = re.compile(r'\btrack (order|delivery)\b', re.IGNORECASE)

keurig_tracking_pattern = re.compile(r'Tracking\s*#\s*:\s*(\S+)', re.IGNORECASE)

mail = oop.Email(emailAdress=email_address, emailPassword=email_password)

result = mail.get_email_ids('INBOX', 'ebay@ebay.com', 'keurig@em.keurig.com')

# print(type(result[0]))

# for i in range(len(result)):
#     email = mail.get_html(email_id=result[i])

#     soup = oop.EmailParser(email=email)
#     href = soup.find_pattern('a', track_order_link_pattern)
#     print(href, flush=True)
#     # tracking = soup.find_pattern('td', pattern=keurig_tracking_pattern)
#     # # print(tracking, flush=True)
#     # shipping = soup.get_shipping_address()
#     # print(type(shipping), flush=True)

#     print(soup.get_name(), flush=True)
#     print(soup.get_zip(), flush=True)


file = oop.File(path=tsv_file_path, type='txt')
excel_sheet = oop.File(path=error_excel_path, type='xlsx')

data_to_add = []


    
excel_sheet.append_data(data_to_add)
excel_sheet.save()

    