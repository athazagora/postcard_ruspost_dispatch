import requests
import json

REQ_URL = 'https://www.pochta.ru/api/nano-apps/api/v1/tracking.get-by-barcodes?language=ru'
HEADERS = {
        'Host': 'www.pochta.ru',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:107.0) Gecko/20100101 Firefox/107.0',
        'Accept': 'application/json',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://www.pochta.ru/tracking',
        'Content-Length': '18',
        'Origin': 'https://www.pochta.ru',
        'DNT': '1',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'no-cors',
        'Sec-Fetch-Site': 'same-origin',
        'TE': 'trailers',
        'Content-Type': 'application/json',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    
error_barcodes = []

def get_full_name(resp: str) -> str:
    resp = json.loads(resp)
    recipient = resp['detailedTrackings'][0]['trackingItem']['recipient']
    full_name = [part_name.capitalize() for part_name in recipient.split()[:3]]
    return ' '.join(full_name)


def get_surname(resp: str) -> str:
    resp = json.loads(resp)
    recipient = resp['response'][0]['trackingItem']['recipient']
    return recipient.split()[0].capitalize()

  

def expract_name_by_barcode (barcode) :
    payload = f'["{barcode}"]'
    resp_raw = requests.post(REQ_URL, data=payload, headers=HEADERS)
    
    try:
        full_name = get_full_name(resp_raw.text)
    except:
        error_barcodes.append(barcode)
    return (full_name)

