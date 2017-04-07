#!/usr/local/bin/python
import requests
from lxml import etree
import sys
import os

def ensure_dir(file_path):
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '*'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    sys.stdout.write('\r')
    sys.stdout.write('%s |%s| %s%% %s' % (prefix, bar, percent, suffix))
    sys.stdout.flush()
    # Print New Line on Complete
    if iteration == total:
        print()

cookie = raw_input(">>> Cookie: ")
total_pages = int(raw_input(">>> Total Pages: "))
folder = raw_input(">>> Download Folder: ") + '/'
begin = raw_input(">>> Begin Date: ")
end = raw_input(">>> End Date: ")

url = 'https://www.einvoice.nat.gov.tw/APMEMBERVAN/Invoice/InvStoreByBuyer!queryInvoice'
download_url = 'https://www.einvoice.nat.gov.tw/APMEMBERVAN/Invoice/InvStoreByBuyer!downloadInvStoreListExcel'
headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-US,en;q=0.8,zh-CN;q=0.6,zh-TW;q=0.4',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Cookie': cookie,
    'Host': 'www.einvoice.nat.gov.tw',
    'Origin': 'https://www.einvoice.nat.gov.tw',
    'Referer': 'https://www.einvoice.nat.gov.tw/APMEMBERVAN/Invoice/InvStoreByBuyer',
    'Upgrade-Insecure-Requests': 1,
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'
}
payload = {
    'invNoQueryRange': 1000,
    'invDateQueryRange': 1,
    'currentPage': 1,
    'pageSize': 40,
    'invoiceQueryVO.invDateBegin': begin,
    'invoiceQueryVO.invDateEnd': end,
    'invoiceQueryVO.showCarriertype': 0,
    'queryCarrierType': 1,
    'queryCarrierTypeTemp': 1
}

row = 1
try:
    printProgressBar(0, total_pages, prefix = 'Progress:', suffix = 'Complete', length = 50)
    for page in range(1, total_pages + 1):
        payload['currentPage'] = page
        r = requests.post(url, headers=headers, data=payload)
        html = etree.HTML(r.text)

        values = html.xpath("//*/input[@id='__checkbox_InvoiceForm_chk']/@value")
        payload['chk'] = values
        payload['__checkbox_chk'] = values
        headers['Referer'] = url
        download = requests.post(download_url, headers=headers, data=payload)

        ensure_dir(folder)
        with open(folder + 'ExcelReport_' + str(page) + '.csv', 'wb') as file:
            file.write(download.content)

        printProgressBar(page, total_pages, prefix = 'Progress:', suffix = 'Complete', length = 50)
except Exception as e:
    print("exception while parsing invoices: " + str(e))
