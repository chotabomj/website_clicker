import json
import os
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def get_tuple_to_lowercase (input_tuple : tuple):   
    result_tuple = tuple(item.lower() if type(item) is str else item for item in input_tuple) 
    return result_tuple


def tuple_match (A, B):
    if len(A) <= len(B):
        return A == tuple(b for b in B if b in A)
    else:
        return B == tuple(a for a in A if a in B)


def get_row_indexes(sheet, keys_tuple: tuple, start_row=1, start_col=1):
    indexes = {}     
    current_row_index = 1                           
    for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row, min_col=start_col, max_col=sheet.max_column, values_only=True):
        row = get_tuple_to_lowercase(row)
        if tuple_match(keys_tuple, row):
            for key in keys_tuple:
                indexes[key] = row.index(key)
            return current_row_index, indexes
        current_row_index += 1


def create_users_json(sheet, header):
    users = {}
    user = {}
    user_id = 0
    header_row, user_header_indexes = get_row_indexes(sheet, header) 
    start_row = header_row + 1           # первая строчка после шапки    
    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        row = get_tuple_to_lowercase(row)
        for column_name in header:
            user[column_name] = row[user_header_indexes[column_name]]
        user['completed'] = False
        users[user_id] = user.copy()
        user_id += 1 
    return users
    
    
def main():
    # формирование json файла
    CURRENT_DIR = os.path.realpath(os.getcwd())
    json_exists = False
    
    #должно стоять перед циклом
    if not json_exists:
        header = ('логин','пароль','пол', 'возраст')
        FILENAME = '392515-0019-7e.xlsx' # добавить аргумент скрипта | единственный файл 
        workbook = load_workbook(filename=FILENAME, read_only=True, data_only=True)
        sheet = workbook.active
        users_json = create_users_json(sheet, header)
        with open('users.json' , 'w', encoding='utf-8') as outfile:
            json.dump(users_json, outfile)
    
    for File in os.listdir(CURRENT_DIR):
        if File.endswith('.json'):
            with open(File, 'r', encoding='utf-8') as json_file:
                users_json_dict = json.load(json_file)
            json_exists = True 

    # web scraping 
    options = Options()
    options.headless = False
    options.add_argument("--window-size=800,600")

    DRIVER_NAME = 'chromedriver.exe'
    PAGE_NAME = 'http://39.soctest.ru'

    DRIVER_PATH = os.path.realpath(os.path.join(os.getcwd(), DRIVER_NAME))
    driver = webdriver.Chrome(options = options, executable_path=DRIVER_PATH) 
    driver.get(PAGE_NAME)

    login_input = driver.find_element_by_xpath('//*[@id="test_user_login"]')
    login_input.send_keys(users_json_dict['0']['логин'])

    password_input = driver.find_element_by_xpath('//*[@id="test_user_password"]')
    password_input.send_keys(users_json_dict['0']['пароль'])

    driver.quit()   

    
if __name__ == '__main__':
    main()
