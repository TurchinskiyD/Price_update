import requests
import xlrd
import time

start = time.time()
# посилання на файл
url = "https://www.dropbox.com/s/qzg0r695jx3m8iw/rozn.XLS?dl=1"

# виконати запит GET до сервера та отримати відповідь
response = requests.get(url)

# зберегти вміст відповіді в файл
with open("price/shambala.xls", "wb") as f:
    f.write(response.content)

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

print(data_shamb)
print(len(data_shamb))
end = time.time() - start
print(f'Файл "shambala.xls" перезаписано. Час виконання {int(end)} секунд.')
