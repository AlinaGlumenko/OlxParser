import queue
import threading
import requests
import time

import xlsxwriter
from lxml.etree import XMLSyntaxError
from lxml.html import fromstring

URL = 'https://www.olx.ua/nedvizhimost/arenda-kvartir/od/?page='
ITEM_PATH = '.wrap .x-large a'
DESCR_PATH = '#textContent p'
PRICE_PATH = '.price-label strong'
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
COUNTER = 1000


def thread_method(result, result_queue):
    response = requests.get(result['url'], headers=headers)
    details_html = response.content.decode('utf-8')
    try:
        details_doc = fromstring(details_html)
    except XMLSyntaxError:
        return

    for elem in details_doc.cssselect(PRICE_PATH):
        result['price'] = elem.text

    for elem in details_doc.cssselect(DESCR_PATH):
        result['descr'] = ' '.join(elem.text.split())

    result_queue.put(result)


def parse():
    results = []
    page_counter = 1

    while len(results) < COUNTER:
        response = requests.get(URL + str(page_counter), headers=headers)
        list_html = response.content.decode('utf-8')
        list_doc = fromstring(list_html)

        for elem in list_doc.cssselect(ITEM_PATH):
            title = elem.cssselect('strong')[0].text
            url = elem.get('href')

            results.append({'title': title, 'url': url})

        page_counter += 1

    result_queue = queue.Queue()
    threads = [threading.Thread(target=thread_method, args=(result, result_queue)) for result in results]

    for t in threads:
        t.start()
    for t in threads:
        t.join()

    return list(result_queue.queue)


def export_excel(filename, results):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})
    field_names = ('Название', 'URL', 'Цена', 'Описание')
    fields = ('title', 'url', 'price', 'descr')

    for i, field in enumerate(field_names):
        worksheet.write(0, i, field, bold)

        for j, result in enumerate(results):
            worksheet.write(j + 1, i, result[fields[i]])

    workbook.close()


def main():
    start_time = time.time()
    result = parse()
    print("--- %s seconds ---" % (time.time() - start_time))
    print(result)
    print(len(result))
    export_excel('result.xlsx', result)


if __name__ == '__main__':
    main()

