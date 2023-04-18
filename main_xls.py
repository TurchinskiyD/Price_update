import xml.etree.ElementTree as ET, openpyxl, xlrd, urllib.request, time

import requests

start = time.time()

# посилання на файл
url = "https://outfitter.in.ua/content/export/1dc57c3051db9aa40919d7d71ef7b23e.xlsx"

# виконати запит GET до сервера та отримати відповідь
response = requests.get(url)

# зберегти вміст відповіді в файл
with open("price/outfitter.xlsx", "wb") as f:
    f.write(response.content)

# завантажити книгу Excel з файлу
wb = openpyxl.load_workbook(filename="price/outfitter.xlsx")

# отримати активний аркуш
ws = wb.active

# створити порожній словник для зберігання даних
data_outfitter = {}

# прочитати дані з кожного рядка (крім першого, який містить заголовки стовпців)
for row in ws.iter_rows(min_row=2, values_only=True):

    # створити словник з даних рядка
    offer_data = {"available": "Немає в наявності", "price": row[9], "name": row[3]}


    data_outfitter[str(row[0])] = offer_data

# print(data_outfitter)
print(f'Довжина словника data_outfitter - {len(data_outfitter)} елементів.')

data_for_update = {}

tre_atl = ET.parse('price/atlantmarket.xml')
root_atl = tre_atl.getroot()

data_atl = {}

for offer_atl in root_atl.findall('./shop/offers/offer'):
    vendor_code_atl = offer_atl.find('barcode').text

    offer_data_atl = {}

    if offer_atl.get('available') == "true":
        offer_data_atl['available'] = "В наявності"
    elif offer_atl.get('available') == "false":
        offer_data_atl['available'] = 'Немає в наявності'

    offer_data_atl['price'] = float(offer_atl.find('price').text)
    offer_data_atl['name'] = offer_atl.find('name').text

    data_atl[vendor_code_atl] = offer_data_atl
    data_for_update[vendor_code_atl] = offer_data_atl

# print(data_atl)
print(f'Довжина словника data_atl - {len(data_atl)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


tree_adr = ET.parse('price/adrenalin.xml')
root_adr = tree_adr.getroot()

data_adr = {}

for offer_adr in root_adr.findall('./item'):
    vendor_code_adr = offer_adr.find('code').text
    offer_data_adr = {}

    if offer_adr.find('stock').text == "Y":
        offer_data_adr['available'] = "В наявності"
    else:
        offer_data_adr['available'] = 'Немає в наявності'

    # offer_data_adr['price'] = float(offer_adr.find('rrc').text)

    rrc_elem = offer_adr.find('rrc')
    if rrc_elem is not None and rrc_elem.text is not None:
        offer_data_adr['price'] = float(rrc_elem.text)

    data_adr[vendor_code_adr] = offer_data_adr
    data_for_update[vendor_code_adr] = offer_data_adr

print(f'Довжина словника data_adr - {len(data_adr)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


tree_dosp = ET.parse('price/dospehi.xml')
root_dosp = tree_dosp.getroot()

data_dosp = {}

for offer_dosp in root_dosp.findall('./shop/offers/offer'):
    vendor_code_dosp = offer_dosp.find('vendorCode').text
    offer_data_dosp = {}

    if offer_dosp.get('available') == "true":
        offer_data_dosp['available'] = "В наявності"
    else:
        offer_data_dosp['available'] = 'Немає в наявності'

    offer_data_dosp['price'] = float(offer_dosp.find('price').text)
    offer_data_dosp['name'] = offer_dosp.find('name').text

    data_dosp[vendor_code_dosp] = offer_data_dosp
    data_for_update[vendor_code_dosp] = offer_data_dosp

print(f'Довжина словника data_dosp - {len(data_dosp)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


# завантажити книгу Excel з файлу
wb = openpyxl.load_workbook(filename="price/swa.xlsx")

# отримати активний аркуш
ws = wb.active

# створити порожній словник для зберігання даних
data_swa = {}

# прочитати дані з кожного рядка (крім першого, який містить заголовки стовпців)
for row in ws.iter_rows(min_row=2, values_only=True):
    # створити словник з даних рядка
    item_data = {"available": row[13], "price": row[7]}

    # додати словник до словника даних, використовуючи артикул як ключ
    data_swa[str(row[0])] = item_data
    data_for_update[str(row[0])] = item_data

print(f'Довжина словника data_swa - {len(data_swa)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


# завантажити книгу Excel з файлу
wb = openpyxl.load_workbook(filename="price/tramp.xlsx")

# отримати активний аркуш
ws = wb.active

# створити порожній словник для зберігання даних
data_trp = {}

# прочитати дані з кожного рядка (крім першого, який містить заголовки стовпців)
for row in ws.iter_rows(min_row=2, values_only=True):
    # створити словник з даних рядка
    __ex_rate = 42
    price = float(row[9]) * __ex_rate
    item_data = {"available": row[12], "price": price}

    # додати словник до словника даних, використовуючи артикул як ключ
    data_trp[str(row[0])] = item_data
    data_for_update[str(row[0])] = item_data

print(f'Довжина словника data_trp - {len(data_trp)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


# завантажити книгу Excel з файлу
workbook = xlrd.open_workbook("price/norfin.xls")

# отримати активний аркуш
worksheet = workbook.sheet_by_index(0)

# створити порожній словник для зберігання даних
data_norf = {}

# прочитати дані з кожного рядка (крім першого, який містить заголовки стовпців)
for row in range(16, worksheet.nrows):
    # створити словник з даних рядка

    if worksheet.cell_value(row, 3) == "Да" or worksheet.cell_value(row, 3) == 'Нет':
        available = 'В наявності'
    else:
        available = 'Немає в наявності'

    item_data = {"available": available, "price": worksheet.cell_value(row, 9)}

    data_norf[str(worksheet.cell_value(row, 1))] = item_data
    data_for_update[str(worksheet.cell_value(row, 1))] = item_data

print(f'Довжина словника data_norf - {len(data_norf)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


# завантажити книгу Excel з файлу
workbook = xlrd.open_workbook("price/shambala.xls")

# отримати активний аркуш
worksheet = workbook.sheet_by_index(0)

# створити порожній словник для зберігання даних
data_shamb = {}

# прочитати дані з кожного рядка (крім перших 6, який містить заголовки стовпців)
for row in range(6, worksheet.nrows):
    # отримати ключ (артикул товару) з комірки таблиці
    key = str(worksheet.cell_value(row, 4))

    if key != "":

        if worksheet.cell_value(row, 5) or worksheet.cell_value(row, 6) or worksheet.cell_value(row, 7):
            available = 'В наявності'
        else:
            available = 'Немає в наявності'

        if worksheet.cell_value(row, 10) != '':
            price = float(worksheet.cell_value(row, 10))
        else:
            price = 0.0

        # створити словник з даних рядка
        item_data = {"available": available, "price": price}

        data_shamb[key] = item_data
    data_for_update[str(worksheet.cell_value(row, 4))] = item_data

print(f'Довжина словника data_shamb - {len(data_shamb)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


work_book = xlrd.open_workbook("price/kemping.xls")

# отримати активний аркуш
work_sheet = work_book.sheet_by_index(0)

# створити порожній словник для зберігання даних
data_kemp = {}

# прочитати дані з кожного рядка (крім перших пяти, який містить заголовки стовпців)
for row in range(5, work_sheet.nrows):
    # отримати ключ (артикул товару) з комірки таблиці
    key = work_sheet.cell_value(row, 0)
    if key != '':

        # створити словник з даних рядка
        if int(work_sheet.cell_value(row, 6)) > 1:
            available = 'В наявності'
        else:
            available = 'Немає в наявності'

        item_data = {"available": available, "price": work_sheet.cell_value(row, 9)}

        data_kemp[str(int(key))] = item_data
        data_for_update[str(int(key))] = item_data

print(f'Довжина словника data_kemp - {len(data_kemp)} елементів.')
print(f'Довжина словника data_for_update - {len(data_for_update)} елементів.')


# записуемо дані зі словника в якому збирали всі прайси з каталогом магазину outfitter

for key in data_for_update:
    if key in data_outfitter:
        data_outfitter[key]['available'] = data_for_update[key]['available']
        data_outfitter[key]['price'] = data_for_update[key]['price']


workbook = openpyxl.Workbook()

# вибираємо активний лист
worksheet = workbook.active

# додаємо заголовки стовпців
worksheet['A1'] = 'Артикул'
worksheet['B1'] = 'Наявність'
worksheet['C1'] = 'Ціна'
worksheet['D1'] = 'Назва'

keys = list(data_outfitter.keys())

for i in range(len(keys)):
    key = keys[i]

    row_num = i + 2  # рядок для запису даних, починаємо з другого рядка
    worksheet.cell(row=row_num, column=1, value=key)
    worksheet.cell(row=row_num, column=2, value=data_outfitter[key]['available'])
    worksheet.cell(row=row_num, column=3, value=data_outfitter[key]['price'])
    worksheet.cell(row=row_num, column=4, value=data_outfitter[key]['name'])

workbook.save("price_update_xlsx.xlsx")

end = time.time() - start
print(f'Файл оновлення price_update.xlsx сформовано. Час виконання {int(end)} секунд')