import xml.etree.ElementTree as ET
import urllib.request
import time

start = time.time()
url_adr = 'https://opt.adrenalin.od.ua/Adrenalin_products.xml'
urllib.request.urlretrieve(url_adr, 'price/adrenalin.xml')


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


end = time.time() - start
print(data_adr)
print(len(data_adr))
print(f'Файл "adrenalin.xml" перезаписано. Час виконання {int(end)} секунд.')


