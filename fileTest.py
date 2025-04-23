import requests, bs4, re #type:ignore

url = 'https://www.ebay.com/signin/g/v%5E1.1%23i%5E1%23r%5E1%23I%5E3%23p%5E3%23f%5E0%23t%5EUl4xMF8yOkNFRDg1OUZEMTU0QTEwRjA3Nzg1QjQ5QUI0MEY2Q0M0XzFfMSNFXjI2MA%3D%3D?mkevt=1&mkpid=0&emsid=e11401.m147516.l152615&mkcid=7&ch=osgood&euid=3173fd04c29545a780c41ce7b3ff923f&bu=45860262360&exe=0&ext=0&osub=-1%7E1&crd=20250423105120&segname=11401'

headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.10 Safari/605.1.1",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Cache-Control": "max-age=0"
}

tracking = []

try:
    rspn = requests.get(url=url, headers=headers, allow_redirects=True)
    ebay_soup = bs4.BeautifulSoup(rspn.text, 'html.parser')

    found_ = ebay_soup.find('span', string=lambda t: t and 'Number' in t)

    parent_dt = found_.find_parent()
    parent_div = parent_dt.find_parent()

    spans = parent_div.find_all('span')
    tracking.append(spans[-1].get_text())

    print(tracking)
    
except requests.exceptions.RequestException as e:
    print(f"Request failed: {e}")
