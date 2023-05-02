import os, main_adr, main_atl, main_dosp, main_kemp, main_norf, main_shamb, main_swa, main_trp, main_xls, requests


def download_file(url, name):
    file_path = os.path.join("price/", name)

    # виконати запит GET до сервера та отримати відповідь
    try:
        response = requests.get(url)
        response.raise_for_status()

        # зберегти вміст відповіді в файл
        with open(file_path, 'wb') as file:
            file.write(response.content)

        print(f'Файл {name} було успішно завантажено та перезаписано.')

    except requests.exceptions.RequestException as e:
        print(f'Виникла помилка під час завантаження {name}: {str(e)}')


data_for_update = main_xls.data_outfitter


def add_for_update(name, dictionary):
    dict_temp = dictionary
    counter = 0
    for key, value in dict_temp.items():
        if key in data_for_update:
            data_for_update[key]['available'] = value['available']
            data_for_update[key]['price'] = value['price']
            counter += 1
    print(f"В каталозі аутфіттер оновлено {counter} товарів. "
          f"Загльна кількість товарів в прайсі {name} {len(dictionary)} елементів.")


if __name__ == "__main__":

    link_list = [('https://outfitter.in.ua/content/export/1dc57c3051db9aa40919d7d71ef7b23e.xlsx', 'outfitter.xlsx'),
                 ('https://www.dropbox.com/s/qzg0r695jx3m8iw/rozn.XLS?dl=1', 'shambala.xls'),
                 ('https://tramp.ua/content/export/wholesale/tramp.ua_8f14e45fceea167a5a36dedd4bea2543.xlsx',
                  'tramp.xlsx'),
                 ('https://uabest.com.ua/content/export/36.xml', 'dospehi.xml'),
                 ('https://atlantmarket.com.ua/price1/prom/atlantmarketprom(false).xml', 'atlantmarket.xml'),
                 ('https://opt.adrenalin.od.ua/Adrenalin_stock.xml', 'adrenalin.xml'),
                 ('https://sva.com.ua/content/export/80.xlsx', 'swa.xlsx')
                 ]

    for url_link, name_link in link_list:
        download_file(url_link, name_link)

    print(f'Кількість товарів в каталозі аутфіттер {len(data_for_update)}')

    functions_list = [("Кемпінг", main_kemp.kemping_file_operation),
                      ("Шамбала", main_shamb.shamb_file_operation),
                      ("Трамп", main_trp.trp_file_operation),
                      ("Сва", main_swa.swa_file_operation),
                      ("Доспехи", main_dosp.dosp_file_operation),
                      ("Норфін", main_norf.norf_file_operation),
                      ("Атлант", main_atl.atl_file_operation),
                      ("Адреналін", main_adr.adr_file_operation)]

    for name_func, func in functions_list:
        result_dict = func()
        add_for_update(name_func, result_dict)

# print(data_for_update)
