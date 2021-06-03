from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
import os.path


def simple_get(url):
    """
    Attempts to get the content at `url` by making an HTTP GET request.
    If the content-type of response is some kind of HTML/XML, return the
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    """
    Returns True if the response seems to be HTML, False otherwise.
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200
            and content_type is not None
            and content_type.find('html') > -1)


def log_error(e):
    """
    It is always a good idea to log errors.
    This function just prints them, but you can
    make it do anything.
    """
    print(e)


def getPrice(symbol):
    raw_html = simple_get(f'https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={symbol}')
    html = BeautifulSoup(raw_html, 'html.parser')
    ul = html.select('#MainContent_PanelInfo > div.main-indicators > ul.top-list > li:nth-child(1) > span')
    price = str(ul[0]).replace("<span>", "")
    price = price.replace("</span>", "")
    price = price.replace(",", ".")
    return float(price)


symbols = ["AFT", "IPV", "TTA", "YAY", "TTE"]
prices = [0] * 4
cells = ['H4', 'H5', 'H6', 'H7', 'H8']
if os.path.exists('./Portfolio.xlsx'):
    wb = load_workbook('./Portfolio.xlsx')
else:
    wb = Workbook()
ws = wb.active
for index in range(len(symbols)):
    ws[cells[index]] = getPrice(symbols[index])
wb.save('./Portfolio.xlsx')
