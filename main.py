from openpyxl import load_workbook
import os
from openpyxl_image_loader import SheetImageLoader
import requests
import cfscrape


book = load_workbook(filename=os.path.abspath("prais.xlsx"))

sheet = book['П Р А Й С (ПТО)']
image_loader = SheetImageLoader(sheet)



#48 1118
for i in range(48, 49):
    i = str(i)
    product_code = sheet['A' + i].value
    name = sheet['B' + i].value
    age = sheet['C' + i].value

    try:
        image = image_loader.get('D' + i)
    except:
        image = None
        print(f'В клетке D{i} не найдено картинки')

    height = sheet['E' + i].value
    length = sheet['F' + i].value
    width = sheet['G' + i].value
    params = sheet['H' + i].value
    weight = sheet['I' + i].value
    concrete = sheet['J' + i].value
    installation_time = sheet['K' + i].value
    price = sheet['L' + i].value
    age_start = None
    age_end = None


    if age:
        integers = []
        n = 0
        while n < len(age):
            s_int = ''
            while n < len(age) and '0' <= age[n] <= '9':
                s_int += age[n]
                n += 1
            n += 1
            if s_int != '':
                integers.append(int(s_int))

        age_start = min(integers)

        if len(integers) > 1:
            age_end = max(integers)


    form_data = {
        'name': name,
        'product_code': product_code,
        'price': price,
    }

    if height:
        form_data['height'] = height
    if length:
        form_data['length'] = length
    if width:
        form_data['width'] = width
    if weight:
        form_data['weight'] = weight
    if params:
        form_data['params'] = params
    if concrete:
        form_data['concrete'] = concrete
    if installation_time:
        form_data['installation_time'] = installation_time
    if age_start:
        form_data['age_start'] = age_start
    if age_end:
        form_data['age_end'] = age_end


    def get_session():
        session = requests.Session()
        session.headers = {
        'Host':'www.artstation.com',
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:69.0)   Gecko/20100101 Firefox/69.0',
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language':'ru,en-US;q=0.5',
        'Accept-Encoding':'gzip, deflate, br',
        'DNT':'1',
        'Connection':'keep-alive',
        'Upgrade-Insecure-Requests':'1',
        'Pragma':'no-cache',
        'Cache-Control':'no-cache'}
        return cfscrape.create_scraper(sess=session)


    if image:
        image.save('D:/Prog/Projects/Parser/image.png')
        img = open('image.png', 'rb')
        files = {'photo': img}
        session = get_session()
        p = session.post('http://127.0.0.1:8000/admin_panel/products/add', data=form_data)
        img.close()
        os.remove('D:/Prog/Projects/Parser/image.png')
    else:
        session = get_session()
        p = session.post('http://127.0.0.1:8000/admin_panel/products/add', data=form_data)

