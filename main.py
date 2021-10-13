import json
from openpyxl import load_workbook

# function restrictions?
def get_tuple_to_lowercase (input_tuple : tuple):   
    result_tuple = tuple(item.lower() if type(item) is str else item for item in input_tuple) 
    return result_tuple


def tuple_match (A, B):
    if len(A) <= len(B):
        return A == tuple(b for b in B if b in A)
    else:
        return B == tuple(a for a in A if a in B)


def get_header_indexes(sheet, keys_tuple: tuple, start_row=1, start_col=1):
    indexes = {}     
    current_row_index = 1                           
    for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row, min_col=start_col, max_col=sheet.max_column, values_only=True):
        row = get_tuple_to_lowercase(row)
        if tuple_match(keys_tuple, row):
            for key in keys_tuple:
                indexes[key] = row.index(key)
            return indexes, current_row_index
        current_row_index += 1


def create_users_json(sheet, user_indexes, header_row):
    users = {}
    start_row = header_row + 1           # первая строчка после шапки
    user_index = user_indexes['login_index']
    user_password = user_indexes['password_index']
    user_sex = user_indexes['sex_index']
    user_age = user_indexes['age_index']
    user_id = 0
    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        user = {
            'completed_test': False,
            'login': row[user_index],
            'password': row[user_password],
            'sex': row[user_sex],
            'age': row[user_age]
        }
        users[user_id] = user
        user_id += 1
    return json.dumps(users, indent=4, ensure_ascii=False) # ensure_ascii=False подключает utf-8
    

def main():
    
    keys_to_find = ('логин','пароль','пол','возраст')
    
    # filename in another directory?
    FILENAME = '392515-0019-7e.xlsx' # добавить аргумент скрипта | единственный файл 
    workbook = load_workbook(filename=FILENAME, read_only=True, data_only=True)
    sheet = workbook.active

    # сделать так, чтобы получало по списку значений вместо железных значений в функции
    user_indexes, header_index = get_header_indexes(sheet, keys_to_find) 
    print(user_indexes)
    # users_json = create_users_json(sheet, user_indexes)
    # print(users_json)


    # users = json.loads(users_json)
    # print(users['0']['login'])
    # print(users)


if __name__ == '__main__':
    main()
