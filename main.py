from openpyxl import load_workbook
import os
from openpyxl_image_loader import SheetImageLoader


book = load_workbook(filename=os.path.abspath("prais.xlsx"))

sheet = book['П Р А Й С (ПТО)']
image_loader = SheetImageLoader(sheet)

#48 1118
for i in range(48, 1118):
    i = str(i)
    product_code = sheet['A' + i].value
    name = sheet['B' + i].value
    age = sheet['C' + i].value
    image = image_loader.get('D' + i)
    height = int(sheet['E' + i].value)
    length = int(sheet['F' + i].value)
    width = int(sheet['G' + i].value)
    params = sheet['H' + i].value
    weight = float(sheet['I' + i].value)
    concrete = float(sheet['J' + i].value)
    installation_time = float(sheet['K' + i].value)
    price = float(sheet['L' + i].value)


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
