from openpyxl import load_workbook
import os
from openpyxl_image_loader import SheetImageLoader
import requests


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
        'price': float(price),
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


    if image:
        image.save('D:/Prog/Projects/Parser/image.png')
        img = open('image.png', 'rb')
        files = {'photo': img}
    else:
        files = {}

    with requests.Session() as s:
        p = s.post('http://127.0.0.1:8000/admin_panel/products/add', data=form_data, files=files)

    if image:
        img.close()
        os.remove('D:/Prog/Projects/Parser/image.png')
