import email_handling, requests, error, logging  #type:ignore
import pandas as pd #type:ignore

logger = logging.getLogger(__name__)

# To be moved to file email_utils
def get_carrier(tracking_number):
    try:
        if tracking_number.startswith('1Z'):
            return 'UPS'
        elif len(tracking_number) in (15, 12):
            return 'FedEx'
        elif tracking_number.startswith('92') or tracking_number.startswith('94'):
            return 'USPS'
    except Exception as e:
        print(f'Error: {e}')
        logger.exception("Error: during get_carrier()")
    
    
def get_backup_tracking(url):
    try:
        tracking = None
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.10 Safari/605.1.1",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "en-US,en;q=0.5",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Cache-Control": "max-age=0"
        }
        i = 0
        status = True
        while status:
            rspn = requests.get(url=url, headers=headers, allow_redirects=True)
            
            if rspn.status_code != 200:
                raise error.No_Tracking_Number(f'<html><p>No tracking number found in email <br /><br />P.S. There might be more issues with this email</p><a href="{url}">Track Order</a>')
            
            ebay_soup = email_handling.EmailParser(rspn.text)

            if ebay_soup.find_element('h1', 'Please verify yourself to continue'):
                headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:138.0) Gecko/20100101 Firefox/138.0'
            else:
                tracking = ebay_soup.get_back_up_tracking()

                print(tracking, flush=True)

                if tracking is not None and len(tracking[0]) < 10:
                    raise error.No_Tracking_Number(f'<html><p>No tracking number found in email <br /><br />P.S. There might be more issues with this email</p><a href="{url}">Track Order</a>')
                logger.info('Got tracking from url provided in email')
                return tracking
            if i >= 2 or tracking is not None:
                status = False
            i += 1
    except Exception:
        logger.exception(f'Error: during get_backup_tracking()')
        
def convert_file(file_path, new_path, sheet):
    try:
        file_to_convert = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet)
        file_to_convert.to_csv(new_path, sep='\t')
        logger.info('Succefully converted to .tsv file')
    except Exception:
        logger.exception('Error: during get convert_file()')