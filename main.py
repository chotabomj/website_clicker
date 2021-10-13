import json
from openpyxl import load_workbook

# function restrictions?
def get_tuple_to_lowercase (input_tuple : tuple):   
    result_tuple = tuple(item.lower() if type(item) is str else item for item in input_tuple) 
    return result_tuple


def get_column_indexes(sheet, start_row=1, start_col=1) -> dict:
    indexes = {}
    row_index = 1                                       #openpyxl считает строки с 1
    for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row, min_col=start_col, max_col=sheet.max_column, values_only=True):
        row = get_tuple_to_lowercase(row)
        for item in row:
            if item in ('логин', 'пароль', 'класс', 'фамилия', 'сайт', 'пол', 'дата рождения', 'возраст'):
                indexes['login_index'] = row.index('логин')
                indexes['password_index'] = row.index('пароль')
                indexes['sex_index'] = row.index('пол')
                indexes['date_index'] = row.index('дата рождения')
                indexes['age_index'] = row.index('возраст')
                indexes['row_index'] = row_index        # строчка, на которой нашли шапку
                return indexes
        row_index += 1


def create_users_json(sheet, user_indexes):
    users = {}
    start_row = user_indexes['row_index'] + 1           # первая строчка после шапки
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
    # filename in another directory?
    FILENAME = '392515-0019-7e.xlsx' # добавить аргумент скрипта | единственный файл 
    workbook = load_workbook(filename=FILENAME, read_only=True, data_only=True)
    sheet = workbook.active

    # сделать так, чтобы получало по списку значений вместо железных значений в функции
    user_indexes = get_column_indexes(sheet) 
    users_json = create_users_json(sheet, user_indexes)
    print(users_json)
    # users = json.loads(users_json)
    # print(users['0']['login'])
    # print(users)


if __name__ == '__main__':
    main()
