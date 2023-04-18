import requests
import openpyxl
import time


start = time.time()
# посилання на файл
url = "https://tramp.ua/content/export/wholesale/tramp.ua_8f14e45fceea167a5a36dedd4bea2543.xlsx"

# виконати запит GET до сервера та отримати відповідь
response = requests.get(url)

# зберегти вміст відповіді в файл
with open("price/tramp.xlsx", "wb") as f:
    f.write(response.content)
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


# вивести словник даних
print(data_trp)

end = time.time() - start
print(f'Файл "tramp.xlsx" перезаписано. Час виконання {int(end)} секунд.')