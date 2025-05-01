import requests, error, emailhandling #type:ignore
from bs4 import BeautifulSoup  #type:ignore


headers = {
    # "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.10 Safari/605.1.1",
    "User-Agent": 'PostmanRuntime/7.43.4',
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Cache-Control": "max-age=0"
}

try:
    href = 'https://www.ebay.com/signin/g/v%5E1.1%23i%5E1%23p%5E3%23I%5E3%23r%5E1%23f%5E0%23t%5EUl4xMF8xMDpBQzlFNUFEQTg0N0YzNUFFNkNFREU2MDU0QjY1QjFFRF8xXzEjRV4yNjA%3D?mkevt=1&mkpid=0&emsid=e11401.m147516.l152615&mkcid=7&ch=osgood&euid=1256cf3002484c84947ec49842682a6a&bu=45858376102&osub=-1%7E1&crd=20250501082920&segname=11401'
    tracking = None
    i = 0
    status = True
    while status:
        print(headers)
        rspn = requests.get(url=href, headers=headers, allow_redirects=True)

        if rspn.status_code != 200:
            raise error.No_Tracking_Number(f'<html><p>No tracking number found in email <br /> where customer shipping address is: <br /> and the order number is <br /><br /><br />P.S. There might be more issues with this email</p><a href="{href}">Track Order</a>')
                            
        ebay_soup = emailhandling.EmailParser(rspn.text)

        # print(ebay_soup.prettify())
        
        if ebay_soup.find_element('h1', 'Please verify yourself to continue'):
            print(True)
            headers["User-Agent"] = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.10 Safari/605.1.1'
            i += 1
            # if i > 0:
            #     raise error.No_Tracking_Number('no tracking number')
        else:    
            tracking = ebay_soup.get_back_up_tracking()

            print(tracking, flush=True)
        if i >= 2 or tracking is not None:
            status = False

        i += 1
        tracking = ['',]

        if tracking is not None and len(tracking[0]) < 10 :
            raise error.No_Tracking_Number(f'<html><p>No tracking number found in email <br /> where customer shipping address is: <br /> and the order number is <br /><br /><br />P.S. There might be more issues with this email</p><a href="{href}">Track Order</a>')
except Exception as e:
    print(f'error: {e}')
       