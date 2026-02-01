# Экспортрует архивные заказы из aliexpress.ru с использованием управления браузером FireFox
# Проверялось на python 3.10, убедитесь, что используемые библиотеки установлены
# Сначала получите кук - get_cookies.py, логинитесь (может падать!) -> cookies.pickle
# кук работает какое-то время, потом надо новый получать
# Скрипт завязан на разметку, особо опасные места помечены !!
# Количество заказов - сколько отражается, можно полистать - см константу NUM_SCROLL
# Трэк, полученно/не получено здесь не анализируется - можно доделать
# Если что-то изменится в разметке или новое понадобится - см. режим разработчика в браузере и доку на selenium
# Если года в дате заказа нет, то он ниже устанавливается в 2026
# Переноса по словам в многосточных ясейках нет - можно настроить уже в excell
# Повторно заказы не эксортируются - см saved_orders.json, можно экспртировать то, чего еще не было
# для того чтобы не банили имеется задержка sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
import pickle
import time
import json

import os
# Import the byte stream handler.
from io import BytesIO

import xlsxwriter

FN_SAVE = 'saved_orders.json'

NUM_SCROLL = 1
KEYS_LIST = ['IMAGE','NAME','PRICE','QTY','DATE','STORE','ORDER NO']

saved_orders = []
if os.path.exists(FN_SAVE):
    with open(FN_SAVE, 'r') as json_file:  
        saved_orders = json.load(json_file)  

driver = webdriver.Firefox()
driver.get("https://aliexpress.ru")
cookies = pickle.load(open("cookies.pickle", "rb"))
for cookie in cookies:
    driver.add_cookie(cookie)

def create_worksheet_columns(size):
    A_CHAR = 65
    list_of_columns = []
    for i in range(size):
        list_of_columns.append(chr(A_CHAR))
        A_CHAR += 1

    return list_of_columns

def write_to_xlsx_file(file_name,list_dict_items):

    list_of_columns = create_worksheet_columns(len(KEYS_LIST))
    row_index = 1

    workbook = xlsxwriter.Workbook(file_name)
    workbook.formats[0].set_font_size(12)
    worksheet = workbook.add_worksheet()

    for col in range(0,500):
        # worksheet.set_column(col, col, 70)
        worksheet.set_column(col,col, 15)
        worksheet.set_row(col, 66) 



    for indx,key in enumerate(KEYS_LIST):
        element_pos = str(list_of_columns[indx])+str(row_index)
        worksheet.write(element_pos, key)

    for dic in list_dict_items:
        row_index += 1

        for indx,key in enumerate(KEYS_LIST):
            value=''
            try:
                value = dic[key]
            except Exception:
                pass
            element_pos = str(list_of_columns[indx])+str(row_index)
            if key == KEYS_LIST[0]:
                # Read an image from a remote url.
                #url = value
                image_data = BytesIO(value)
                print(f'Set img [{len(value)}] to {element_pos}')
                # Write the byte stream image to a cell. Note, the filename must be
                # specified. In this case it will be read from url string.
                try:
                    worksheet.insert_image(element_pos, 'img.png', {'image_data': image_data, 'x_offset': 7, 'y_offset': 7})
                except Exception as e:
                    print(e)
                    pass
            else:
                value = value.replace('\n', '\r\n')
                worksheet.write(element_pos, value)
    workbook.close()


def get_item_details(url):    
    print(f'Getting {url} ...')
    res = []
    driver.get(url)
    order_num = 'unknown'
    store = 'unknwon'
    order_date = 'unknwon'
    el_order_num = driver.find_elements(By.CSS_SELECTOR, "[data-testid='orderNumber']")    
    if el_order_num:
        order_num = el_order_num[0].text
        info: str = el_order_num[0].find_element(By.XPATH, '..').find_element(By.XPATH, '..').text
        i = info.find("от ")  # from Date
        if i > 0:
            order_date = info[i+3:].strip()
            if not '202' in order_date:
                order_date += ' 2026'  # current year

    print(f'Order {order_num}')    
    els_seller = driver.find_elements(By.CSS_SELECTOR, "[data-testid='sellerInfoV2']") 
    if els_seller:
        store = els_seller[0].text + '\n' + els_seller[0].get_attribute('href')

    els_item_img = driver.find_elements(By.CSS_SELECTOR, "[data-testid='product']")    
    item_imgs = []
    for el in els_item_img:
        item_imgs.append(el.find_element(By.TAG_NAME, 'picture').find_element(By.TAG_NAME, 'img'))    

    els_item_name = driver.find_elements(By.CSS_SELECTOR, "[data-testid='productText']")
    if not els_item_name:
        return res
    
    for i, el_name in enumerate(els_item_name):
        
        row_dict = dict()        
        row_dict['IMAGE'] = item_imgs[i].screenshot_as_png
        el_parent = el_name.find_element(By.XPATH, '..')
        # product name + product type + url
        row_dict['NAME'] = el_parent.text + '\n' + el_name.get_attribute('href')
        # !!!
        els_price = el_parent.find_element(By.XPATH, '..').find_elements(By.CSS_SELECTOR, 'div.RedOrderDetailsProductsV2_Product__priceDesktop__1nlc3')
        if els_price:
            s = els_price[0].text.split('\n')
            row_dict['PRICE'] = s[0].strip()
            if len(s) > 1:
                row_dict['QTY'] = s[1].strip()
        row_dict['DATE'] = order_date
        row_dict['STORE'] = store
        row_dict['ORDER NO'] = order_num
           
        res.append(row_dict)
        #print(row_dict)
    
    return res


def get_list_of_item_view_details(url):
    print(f'Getting {url} ...')
    driver.get(url)
    links = set()
    items_dict_list = []
    print(f'Scroll down by {NUM_SCROLL} pages...')
    for i in range(NUM_SCROLL):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)  # time for loading
    print(f'Collecting links...')
    view_details = driver.find_elements(By.CSS_SELECTOR, 'a.RedOrderList_OrderItem__link__1tjf5') # !!!
    # Fill list of links to check
    for indx,view in enumerate(view_details):
        link = view.get_attribute('href')
        links.add(link)
    
    print(f'{len(links)} links collected')
    links = sorted(links, reverse=True)     # Most recent first

    # Get dict representing a single ROW
    for link in links:
        p = link.split('/')     #https://aliexpress.ru/order-list/<number>?filterName=archive
        s = p[-1]
        p = s.split('?')
        order_id = p[0]
        if order_id in saved_orders:    # exported allready?
            continue
        time.sleep(3)  # prevent ddos filter
        items = get_item_details(link)
        items_dict_list.extend(items)
        saved_orders.append(order_id)

    driver.close()
    return items_dict_list


if __name__ == '__main__':
    items_dict_list = get_list_of_item_view_details('https://aliexpress.ru/order-list?filterName=archive')
    write_to_xlsx_file('aliexpress_orders.xlsx',items_dict_list)
    with open(FN_SAVE, "w") as f:
        json.dump(saved_orders, f)

