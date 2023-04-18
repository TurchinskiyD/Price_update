import requests
import openpyxl

# посилання на файл
url = "https://sva.com.ua/content/export/80.xlsx"

# виконати запит GET до сервера та отримати відповідь
response = requests.get(url)

# зберегти вміст відповіді в файл
with open("price/swa.xlsx", "wb") as f:
    f.write(response.content)

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

# вивести словник даних
print(data_swa)
print(len(data_swa))
