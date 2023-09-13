import requests
import openaiAPI
import hashlib
from datetime import datetime
import gspread
import math
import re


def make_headers_dict_from_sheet_attributes(sheets, headers_dicts):
    for i in range(len((sheets))):
        sheet = sheets[i]
        headers_dict = headers_dicts[i]
        for elem in enumerate(sheet[0]):
            headers_dict[elem[1]] = elem[0]


def get_request_token(base_url):
    url = f'{base_url}/oauth/requesttoken'
    responce = requests.get(url)
    data = responce.json()
    return data['RequestToken']


def get_access_token(base_url, req_token, pub_key, hash_token):
    url = f'{base_url}/oauth/accesstoken'
    params = {
        'oauth_token': req_token,
        'grant_type': 'api',
        'username': pub_key,
        'password': hash_token
    }
    responce = requests.get(url, params=params)

    data = responce.json()
    return data['AccessToken']


def get_users_list(acc_token, base_url, datetime_from, date_time_to):
    url = f'{base_url}/users'
    cur_datetime = datetime.now().strftime('%Y-%m-%dT%H:%M')

    params = {
        'oauth_token': acc_token,
        'dateFrom': datetime_from,
        'dateTo': date_time_to
    }
    responce = requests.get(url, params=params)
    data = responce.json()
    return data['Result']


def get_user_by_id(acc_token, id):
    url = f'{base_url}/users/{id}'
    params = {
        'oauth_token': acc_token
    }
    responce = requests.get(url, params)
    data = responce.json()
    return data['Result'][0]


def get_orders_list(acc_token, dateFrom, dateTo):
    url = f'{base_url}/orders/betweenDates'
    params = {
        'oauth_token': acc_token,
        'dateFrom': dateFrom,
        'dateTo': dateTo
    }
    responce = requests.get(url, params)
    data = responce.json()
    return data['Result']


def get_order_by_id(acc_token, base_url, order_id):
    url = f'{base_url}/orders/{order_id}'
    params = {
        'oauth_token': acc_token
    }
    responce = requests.get(url, params)
    data = responce.json()
    return data['Result'][0]


def get_order_ids(orders):
    order_ids = []
    for order in orders:
        order_ids.append(order['Id'])
    return order_ids


def get_order_items(acc_token, base_url, order_ids):  # order_ids is list of str ids
    if type(order_ids) == list:
        ids = ','.join(map(str, order_ids))
        url = f'{base_url}/orders/items'
        params = {
            'oauth_token': acc_token,
            'orderIds': ids
        }
        responce = requests.post(url, params)
        data = responce.json()
        return data['Result']
    print('order_ids must be list type')


def get_prod_categories(acc_token, base_url):
    url = f'{base_url}/products'
    params = {
        'oauth_token': acc_token,
    }
    responce = requests.get(url, params)
    data = responce.json()
    return data['Result']


def get_prods_from_category(acc_token, base_url, cat_id):
    url = f'{base_url}/products/{cat_id}'
    params = {
        'oauth_token': acc_token,
        'categoryId': int(cat_id)
    }
    responce = requests.get(url, params)
    data = responce.json()
    return data['Result']


def find_lin_first_empty_cell_in_col(sheet, column):
    for row in range(0, len(sheet)):
        if sheet[row][column] == '' or sheet[row][column] == None:
            return row

    return row + 1


# need to rework for gspread
# def find_bin_first_empty_cell_in_col(sheet, column):
#     l = 1
#     r = 10000
#     while l < r:
#         m = (l + r)//2
#         if sheet.cell(row=m, column=column).value == None:
#             r = m
#         else:
#             l = m + 1
#     return l

def maxrange(max_len_row, num_row):  # твои атрибуты (max_col, MAXR)
    if max_len_row <= 26:
        return f"A1:{chr(ord('A') + max_len_row - 1)}{num_row}"
    else:
        max_len_row_new = max_len_row - 26
        return "A1:" + chr(ord('A') + max_len_row // 27 - 1) + chr(ord('A') + max_len_row_new - 1) + str(num_row)


def number_to_column_name(n):
    result = ""
    while n > 0:
        remainder = (n - 1) % 26
        char = chr(ord('A') + remainder)
        result = char + result
        n = (n - 1) // 26
    return result


def write_all_user_to_excel(spreadsheet, users_sheet, users_gsheets_headers, users_api, first_empty_user):
    users_id_skip = [143306366, 147071135, 150060124, 152689163, 161298395]
    max_col = len(users_gsheets_headers)
    row_to_write = first_empty_user + 1
    users_worksheet = spreadsheet.worksheet("Clients")

    if row_to_write > len(users_sheet):
        users_sheet.append([''])
    for row in range(row_to_write, len(users_api) + row_to_write):
        user = users_api[row - row_to_write]
        try:
            for key in user:
                if key in users_gsheets_headers:
                    users_sheet[row][users_gsheets_headers[key]] = user[key]
        except:
            new_row = [0] * max_col
            for key in user:
                if key in users_gsheets_headers:
                    cell_value = user[key]
                    new_row[users_gsheets_headers[key]] = cell_value
            users_sheet.append(new_row)
    range_to_update = maxrange(max_col, row + 1)

    try:
        users_worksheet.update(range_to_update, users_sheet, value_input_option='USER_ENTERED')
    except gspread.exceptions.APIError as e:
        print(f"An API error occurred in users update: {e}")


def time_to_finish(milliseconds, hours):
    finish = milliseconds / 1000 + int(hours) * 3600
    normal_date = datetime.fromtimestamp(finish)
    formatted_date = normal_date.strftime('%d.%m.%Y %H:%M')
    return formatted_date


def write_all_order_to_excel(spreadsheet, orders_sheet, orders_gsheets_headers, orders_items, id_order_dict, users_api,
                             first_empty_order):
    name_number_dict = dict()
    banned_users = [143306366, 147071135, 150060124, 152689163, 161298395]
    for user in users_api:
        name_number_dict[int(user['Id'])] = f'{user["FirstName"]} +{user["Phone"]}'

    items_count = len(orders_items)
    max_col = len(orders_gsheets_headers)
    orders_worksheet = spreadsheet.worksheet("Orders")
    row_to_write = first_empty_order + 1
    if row_to_write > len(orders_sheet):
        orders_sheet.append([''])
    for row in range(row_to_write, items_count + row_to_write):
        item = orders_items[row - row_to_write]
        item['Name'] = f'{item["Name"]} - {item["Quantity"]} kom'
        order_id = item['OrderId']
        order = id_order_dict[order_id]

        if order['PaymentSystemUniqueId'] == 'Cash' or order['PaymentStatus'] == 'Paid':
            user_id = int(order['UserId'])
            if user_id in banned_users:
                continue
            try:
                for key in item:
                    if key in orders_gsheets_headers:
                        orders_sheet[row][orders_gsheets_headers[key]] = item[key]

                orders_sheet[row][orders_gsheets_headers['UserId']] = user_id
                orders_sheet[row][orders_gsheets_headers['Custumer_Name_phone_number']] = name_number_dict[user_id]
                orders_sheet[row][orders_gsheets_headers['Status']] = order['Status']
            except:
                new_row = [0] * max_col
                for key in item:
                    if key in orders_gsheets_headers:
                        cell_value = item[key]
                        new_row[orders_gsheets_headers[key]] = cell_value
                new_row[orders_gsheets_headers['UserId']] = user_id
                new_row[orders_gsheets_headers['Custumer_Name_phone_number']] = name_number_dict[user_id]
                new_row[orders_gsheets_headers['Status']] = order['Status']
                orders_sheet.append(new_row)


    range_to_update = maxrange(max_col, row + 1)

    try:
        orders_worksheet.update(range_to_update, orders_sheet, value_input_option='USER_ENTERED')
    except gspread.exceptions.APIError as e:
        print(f"An API error occurred in orders update: {e}")


def write_photobooks_to_excel(item, photobook_gsheet, photobook_gsheet_headers, prod_cat_dict, id_order_dict, prod_id,
                              order_id):
    order = id_order_dict[int(order_id)]
    order_created = int(order['DateCreated'].strip('/Date()'))
    finish_time = time_to_finish(order_created, 72)
    max_col = len(photobook_gsheet_headers)
    option_headers_dict = dict()
    cur_row = len(photobook_gsheet) + 1
    option_headers_dict['Ne treba'] = 'Laminacija'
    option_headers_dict['Ne'] = 'Dekor'
    option_headers_dict['Klasični'] = 'Uglovi'
    last_process_col = photobook_gsheet_headers['Pakovanje'] + 1
    last_proc_col_let = number_to_column_name(last_process_col)
    option_headers_dict[f'=IF({last_proc_col_let}{cur_row} <> 0; "Done"; "In proc")'] = 'Status'
    deadline_col = photobook_gsheet_headers['Rok izrade'] + 1
    deadline_col_let = number_to_column_name(deadline_col)
    option_headers_dict[
        f'=IF(NOW() < {deadline_col_let}{cur_row}; ROUNDDOWN(({deadline_col_let}{cur_row} - NOW()) * 24); "Istekao")'] = 'Imaš sati'
    option_headers_dict[order_id] = 'ID_order'
    option_headers_dict[finish_time] = 'Rok izrade'
    prod_name = prod_cat_dict[prod_id]['Title'].split(' ', 2)[-1].strip('"')
    option_headers_dict[prod_name] = 'Type'
    # print(prod_name)
    # print(item)
    option_headers_dict[int(item['Id'])] = 'Item_id'
    pages = item['PageCount']
    option_headers_dict[pages] = 'Listovi'
    # print(f'Pages: {pages}')
    count = item['Quantity']

    # print(f'Count: {count} kom')
    try:
        thickness = float(item['EditorOutput']['Thickness']) + 0.6
    except:
        thickness = 0
    for option in item['Options']:
        option_name = option['Title'].split(' (')[0]
        option_value = option['Title'].split(' (')[1][:-1].strip()

        if option_name == 'Predlistovi' and option_value != 'Bez predlistova':
            thickness += 2.4
        if option_name == 'Laminacija korice':
            option_headers_dict[option_value] = 'Laminacija'
        elif option_name == 'Predlistovi':
            option_headers_dict[option_value] = 'Predlistovi'
        elif option_name == 'Uglovi':
            option_headers_dict[option_value] = 'Uglovi'

        # print(f'{option_name}: {option_value}')

    for attrib in item['Attributes']:
        attrib_name = attrib['Title']
        attrib_value = attrib['Value']
        if attrib_name == 'Dimenzije':
            option_headers_dict[attrib_value] = 'Dimenzije'
        elif attrib_name == 'Materijal korice' or attrib_name == 'Boja korice':
            option_headers_dict[attrib_value] = 'Korice'
        elif attrib_name == 'Dekor korice':
            option_headers_dict[attrib_value] = 'Dekor'
        if attrib_value == 'Foto korica':
            thickness -= 1.2
            pages -= 1
        # print(f'{attrib_name}: {attrib_value}')
    # print(f'Thickness: {thickness}')
    option_headers_dict[pages] = 'Listovi'
    option_headers_dict[count] = 'Kolicina'
    thickness = math.ceil(thickness)
    option_headers_dict[f'{thickness} mm'] = 'Debljina'

    new_row = [0] * max_col
    for cell_value in option_headers_dict:
        if option_headers_dict[cell_value] in photobook_gsheet_headers:
            new_row[photobook_gsheet_headers[option_headers_dict[cell_value]]] = cell_value
        else:
            print(f'Find new value: {cell_value}, {option_headers_dict[cell_value]}')
    photobook_gsheet.append(new_row)


def write_photos_to_excel(item, photos_gsheet, photos_gsheet_headers, order_id, id_order_dict):
    order = id_order_dict[int(order_id)]
    cur_row = len(photos_gsheet) + 1
    max_col = len(photos_gsheet_headers)
    order_created = int(order['DateCreated'].strip('/Date()'))
    finish_time = time_to_finish(order_created, 24)
    option_headers_dict = dict()
    last_process_col = photos_gsheet_headers['Pakovanje'] + 1
    last_proc_col_let = number_to_column_name(last_process_col)
    option_headers_dict[f'=IF({last_proc_col_let}{cur_row} <> 0; "Done"; "In proc")'] = 'Status'
    deadline_col = photos_gsheet_headers['Rok izrade'] + 1
    deadline_col_let = number_to_column_name(deadline_col)
    option_headers_dict[
        f'=IF(NOW() < {deadline_col_let}{cur_row}; ROUNDDOWN(({deadline_col_let}{cur_row} - NOW()) * 24); "Istekao")'] = 'Imaš sati'
    option_headers_dict['Lustre'] = 'Papir'
    option_headers_dict[order_id] = 'ID_order'
    prod_name = item['Name'].split(' (')[0]
    option_headers_dict[prod_name] = 'Type'
    # print(prod_name)
    # print(item)
    option_headers_dict[int(item['Id'])] = 'Item_id'
    if '[' in item['Name']:
        ivica = item['Name'].split(' [')[-1].strip(']')
        if 'с пол' in ivica:
            ivica = 'sa ivicama'
        elif 'без пол' in ivica:
            ivica = 'bez ivica'
    else:
        ivica = 'bez ivica'
    option_headers_dict[ivica] = 'Ivica'
    # print(f'Ivica: {ivica}')
    count = item['Quantity']
    if 'kom' in str(count):
        option_headers_dict[f'{count}'] = 'Kolicina'
    else:
        option_headers_dict[f'{count} kom.'] = 'Kolicina'
    # print(f'Count: {count} kom')
    for option in item['Options']:
        option_name = option['Title'].split(' (')[0]
        option_value = option['Title'].split(' (')[1][:-1].strip()
        if option_name == 'Papir':
            option_headers_dict[option_value] = 'Papir'
        elif option_name == 'Rok izrade':

            if option_value == '3 sata':
                finish_time = time_to_finish(order_created, 3)
    option_headers_dict[finish_time] = 'Rok izrade'
    # print(f'{option_name}: {option_value}')
    for attrib in item['Attributes']:
        attrib_name = attrib['Title']
        attrib_value = attrib['Value']
        if attrib_name == 'Dimenzije':
            option_headers_dict[attrib_value] = 'Dimenzije'
        elif attrib_name == 'Broj slike':
            option_headers_dict[attrib_value] = 'Kolicina'
        # print(f'{attrib_name}: {attrib_value}')
    new_row = [0] * max_col
    for cell_value in option_headers_dict:
        if option_headers_dict[cell_value] in photos_gsheet_headers:
            new_row[photos_gsheet_headers[option_headers_dict[cell_value]]] = cell_value
        else:
            print(f'Find new value: {cell_value}, {option_headers_dict[cell_value]}')
    photos_gsheet.append(new_row)


def write_canvas_to_excel(item, canvas_gsheet, canvas_gsheet_headers, order_id, id_order_dict):
    order = id_order_dict[int(order_id)]
    cur_row = len(canvas_gsheet) + 1
    max_col = len(canvas_gsheet_headers)
    order_created = int(order['DateCreated'].strip('/Date()'))
    finish_time = time_to_finish(order_created, 24)

    option_headers_dict = dict()
    option_headers_dict[finish_time] = 'Rok izrade'
    last_process_col = canvas_gsheet_headers['Pakovanje'] + 1
    last_proc_col_let = number_to_column_name(last_process_col)
    option_headers_dict[f'=IF({last_proc_col_let}{cur_row} <> 0; "Done"; "In proc")'] = 'Status'
    deadline_col = canvas_gsheet_headers['Rok izrade'] + 1
    deadline_col_let = number_to_column_name(deadline_col)
    option_headers_dict[
        f'=IF(NOW() < {deadline_col_let}{cur_row}; ROUNDDOWN(({deadline_col_let}{cur_row} - NOW()) * 24); "Istekao")'] = 'Imaš sati'

    option_headers_dict['Da'] = 'Uramljivanje'
    option_headers_dict[order_id] = 'ID_order'
    prod_name = item['Name'].split(' (')[0]
    if 'delova' in prod_name:
        prod_name = item['Name'].split(' (')[1][:28]
    option_headers_dict[prod_name] = 'Type'
    # print(prod_name)
    # print(item)
    option_headers_dict[int(item['Id'])] = 'Item_id'
    count = item['Quantity']
    option_headers_dict[count] = 'Kolicina'
    # print(f'Count: {count} kom')
    for option in item['Options']:
        option_name = option['Title'].split(' (')[0]
        option_value = option['Title'].split(' (')[1][:-1].strip()
        if option_name == 'Lakiranje':
            option_headers_dict[option_value] = 'Lakiranje'
        # print(f'{option_name}: {option_value}')
    for attrib in item['Attributes']:
        attrib_name = attrib['Title']
        if attrib_name == 'Tip':
            attrib_name = 'Dimenzije'
        attrib_value = attrib['Value']
        if attrib_name == 'Dimenzije':
            option_headers_dict[attrib_value] = 'Dimenzije'
        elif attrib_name == 'Uramljivanje':
            option_headers_dict[attrib_value] = 'Uramljivanje'
        # print(f'{attrib_name}: {attrib_value}')
    # print(option_headers_dict)
    # print(canvas_gsheet_headers)
    new_row = [0] * max_col
    for cell_value in option_headers_dict:
        if option_headers_dict[cell_value] in canvas_gsheet_headers:
            new_row[canvas_gsheet_headers[option_headers_dict[cell_value]]] = cell_value
        else:
            print(f'Find new value: {cell_value}, {option_headers_dict[cell_value]}')
    canvas_gsheet.append(new_row)


def write_scetchbooks_to_excel(item, scetchbook_gsheet, scetchbook_gsheet_headers, order_id, id_order_dict):
    order = id_order_dict[int(order_id)]
    cur_row = len(scetchbook_gsheet) + 1
    order_created = int(order['DateCreated'].strip('/Date()'))
    finish_time = time_to_finish(order_created, 24)
    max_col = len(scetchbook_gsheet_headers)

    option_headers_dict = dict()
    option_headers_dict[finish_time] = 'Rok izrade'
    last_process_col = scetchbook_gsheet_headers['Pakovanje'] + 1
    last_proc_col_let = number_to_column_name(last_process_col)
    option_headers_dict[f'=IF({last_proc_col_let}{cur_row} <> 0; "Done"; "In proc")'] = 'Status'
    deadline_col = scetchbook_gsheet_headers['Rok izrade'] + 1
    deadline_col_let = number_to_column_name(deadline_col)
    option_headers_dict[
        f'=IF(NOW() < {deadline_col_let}{cur_row}; ROUNDDOWN(({deadline_col_let}{cur_row} - NOW()) * 24); "Istekao")'] = 'Imaš sati'

    option_headers_dict[order_id] = 'ID_order'
    prod_name = item['Attributes'][1]['Value']
    option_headers_dict[prod_name] = 'Type'
    # print(prod_name)
    # print(item)
    option_headers_dict[int(item['Id'])] = 'Item_id'
    count = item['Quantity']
    option_headers_dict[count] = 'Kolicina'
    # print(f'Count: {count} kom')
    for option in item['Options']:
        option_name = option['Title'].split(' (')[0]
        option_value = option['Title'].split(' (')[1][:-1].strip()
        if option_name == 'Predlistovi':
            option_headers_dict[option_value] = 'Predlistovi'
        # print(f'{option_name}: {option_value}')
    for attrib in item['Attributes']:
        attrib_name = attrib['Title']
        attrib_value = attrib['Value']
        if attrib_name == 'Dimenzije':
            option_headers_dict[attrib_value] = 'Dimenzije'
        # print(f'{attrib_name}: {attrib_value}')
    new_row = [0] * max_col
    for cell_value in option_headers_dict:
        if option_headers_dict[cell_value] in scetchbook_gsheet_headers:
            new_row[scetchbook_gsheet_headers[option_headers_dict[cell_value]]] = cell_value
        else:
            print(f'Find new value: {cell_value}, {option_headers_dict[cell_value]}')
    scetchbook_gsheet.append(new_row)


def sheets_dict_create(spreadsheet):
    sheets = dict()
    worksheets = dict()
    sheet_num = len(spreadsheet.worksheets())
    for i in range(sheet_num):
        sheet = spreadsheet.get_worksheet(i)
        sheet_name = sheet.title
        worksheets[sheet_name] = sheet
        sheets[sheet_name] = sheet.get_all_values()
    return sheets, worksheets


def set_category_status(acc_token, base_url, cat_id, enable_status):
    url = f'{base_url}/categories/{cat_id}/status'

    params = {
        'oauth_token': acc_token,
        'typeId': int(cat_id),
        'isEnabled': enable_status
    }
    responce = requests.post(url, params)
    data = responce.json()
    return data['Result']


def make_product_prod_sheet_dict(categories, prods_list, acc_token, base_url):
    category_prod_sheet_dict = dict()
    photos_cat = [categories[0]['Id'], categories[1]['Id'], categories[4]['Id'], categories[6]['Id']]
    photobooks_cat = [categories[5]['Id'], categories[7]['Id'], categories[8]['Id'], categories[9]['Id']]
    canvas_cat = [categories[2]['Id'], categories[3]['Id']]
    scetchbook_cat = [categories[10]['Id']]
    prods_cats = [photobooks_cat, photos_cat, canvas_cat, scetchbook_cat]
    for i in range(len(prods_list)):
        for cat in prods_cats[i]:
            category_prod_sheet_dict[cat] = prods_list[i]
    prod_cat_dict = dict()
    product_prod_sheet_dict = dict()
    for category in categories:
        if category['IsEnabled']:
            cat_id = category['Id']
            for prod in get_prods_from_category(acc_token, base_url, cat_id):
                product_prod_sheet_dict[int(prod['Id'])] = category_prod_sheet_dict[cat_id]
                prod_cat_dict[int(prod['Id'])] = category
    return product_prod_sheet_dict, prod_cat_dict


def extract_substring(s, pattern):
    match = re.search(pattern, s)
    if match:
        return match.group(1)
    else:
        return "Подстрока не найдена"


def update_prodsheets(sheets, worksheets):
    shets_name = list(worksheets.keys())
    prods_start_index = shets_name.index('Photobook')
    for i in range(prods_start_index, len(sheets)):
        sheet = sheets[shets_name[i]]
        worksheet = worksheets[shets_name[i]]
        max_row = len(sheet)
        max_col = len(sheet[0])
        max_range = maxrange(max_col, max_row)

        try:
            worksheet.update(max_range, sheet, value_input_option='USER_ENTERED')
        except gspread.exceptions.APIError as e:
            print(f"An API error occurred in {shets_name} update: {e}")

# def make_category_prod_sheet_dict(categories, prods_list):
#     category_prod_sheet_dict = dict()
#     photos_cat = [categories[0]['Id'], categories[1]['Id'], categories[4]['Id'], categories[6]['Id']]
#     photobooks_cat = [categories[5]['Id'], categories[7]['Id'], categories[8]['Id'], categories[9]['Id']]
#     canvas_cat = [categories[2]['Id'], categories[3]['Id']]
#     scetchbook_cat = [categories[10]['Id']]
#     prods_cats = [photobooks_cat, photos_cat, canvas_cat, scetchbook_cat]
#
#     for i in range(len(prods_list)):
#         for cat in prods_cats[i]:
#             category_prod_sheet_dict[cat] = prods_list[i]
#     return category_prod_sheet_dict

base_url = 'http://api.pixlpark.com'
pub_key = openaiAPI.pub_key
priv_key = openaiAPI.priv_key
req_token = get_request_token(base_url)
hash_token = hashlib.sha1(f'{req_token}{priv_key}'.encode()).hexdigest()
acc_token = get_access_token(base_url, req_token, pub_key, hash_token)
cur_datetime = datetime.now().strftime('%Y-%m-%dT%H:%M')
first_user_date = '2022-03-01T00:00'

users_api = get_users_list(acc_token, base_url, first_user_date, cur_datetime)
users_api.sort(key=lambda x: x['Id'], reverse=True)
banned_users = [143306366, 147071135, 150060124, 152689163, 161298395]

orders_api = get_orders_list(acc_token, first_user_date, cur_datetime)
orders_api.sort(key=lambda x: x['Id'], reverse=True)
orders_api = [order for order in orders_api if order['Status'] != 'Cancelled']
orders_api_ids = [int(order['Id']) for order in orders_api]

credentials = openaiAPI.credentials
client = gspread.service_account_from_dict(credentials)

spreadsheet = client.open("fotozona")

sheets, worksheets = sheets_dict_create(spreadsheet)

orders_gsheets = sheets['Orders']
users_gsheets = sheets['Clients']
photobook_gsheet = sheets['Photobook']
photos_gsheet = sheets['Photos']
canvas_gsheet = sheets['Canvas']
scetchbook_gsheet = sheets['Scetchbook']

orders_gsheets_headers = dict()
users_gsheets_headers = dict()
photobook_gsheet_headers = dict()
photos_gsheet_headers = dict()
canvas_gsheet_headers = dict()
scetchbook_gsheet_headers = dict()

sheet_list = [orders_gsheets, users_gsheets, photobook_gsheet, photos_gsheet, canvas_gsheet, scetchbook_gsheet]
headers_list = [orders_gsheets_headers, users_gsheets_headers, photobook_gsheet_headers, photos_gsheet_headers,
                canvas_gsheet_headers, scetchbook_gsheet_headers]
prods_list = [photobook_gsheet, photos_gsheet, canvas_gsheet, scetchbook_gsheet]

make_headers_dict_from_sheet_attributes(sheet_list, headers_list)

first_empty_order = find_lin_first_empty_cell_in_col(orders_gsheets, orders_gsheets_headers['OrderId'])
first_empty_user = find_lin_first_empty_cell_in_col(users_gsheets, users_gsheets_headers['Id'])

categories = get_prod_categories(acc_token, base_url)
enable_categories = [category for category in categories if category['IsEnabled']]
photos_cat = [enable_categories[0], enable_categories[1], enable_categories[4], enable_categories[6]]
photobooks_cat = [enable_categories[5], enable_categories[7], enable_categories[8], enable_categories[9]]
canvas_cat = [enable_categories[2], enable_categories[3]]
scetchbook_cat = [enable_categories[10]]

product_prod_sheet_dict, prod_cat_dict = make_product_prod_sheet_dict(enable_categories, prods_list, acc_token,
                                                                      base_url)

orders_ids = get_order_ids(orders_api)
id_order_dict = {int(order['Id']): order for order in orders_api}

order_items = get_order_items(acc_token, base_url, orders_ids)
photobook_items = [int(photobook_gsheet[i][-1]) for i in range(1, len(photobook_gsheet)) if
                   photobook_gsheet[i][-1].isdigit()]
photos_items = [int(photos_gsheet[i][-1]) for i in range(1, len(photos_gsheet)) if photos_gsheet[i][-1].isdigit()]
canvas_items = [int(canvas_gsheet[i][-1]) for i in range(1, len(canvas_gsheet)) if canvas_gsheet[i][-1].isdigit()]
scetchbook_items = [int(scetchbook_gsheet[i][-1]) for i in range(1, len(scetchbook_gsheet)) if
                    scetchbook_gsheet[i][-1].isdigit()]


for item in order_items:

    order_id = int(item['OrderId'])
    banned_statuses = ['Cancelled', 'Delivered']
    if id_order_dict[order_id]['Status'] in banned_statuses:
        continue
    prod_id = int(item['MaterialId'])
    item_id = item['Id']
    print(id_order_dict[order_id])
    print(id_order_dict[order_id]['Status'])
    if item_id not in photobook_items and prod_id in product_prod_sheet_dict and product_prod_sheet_dict[
        prod_id] == photobook_gsheet:
        write_photobooks_to_excel(item, photobook_gsheet, photobook_gsheet_headers, prod_cat_dict, id_order_dict,
                                  prod_id, order_id)
    if item_id not in photos_items and prod_id in product_prod_sheet_dict and product_prod_sheet_dict[
        prod_id] == photos_gsheet:
        write_photos_to_excel(item, photos_gsheet, photos_gsheet_headers, order_id, id_order_dict)
    if item_id not in canvas_items and prod_id in product_prod_sheet_dict and product_prod_sheet_dict[
        prod_id] == canvas_gsheet:
        write_canvas_to_excel(item, canvas_gsheet, canvas_gsheet_headers, order_id, id_order_dict)
    if item_id not in scetchbook_items and prod_id in product_prod_sheet_dict and product_prod_sheet_dict[
        prod_id] == scetchbook_gsheet:
        write_scetchbooks_to_excel(item, scetchbook_gsheet, scetchbook_gsheet_headers, order_id, id_order_dict)

write_all_user_to_excel(spreadsheet, users_gsheets, users_gsheets_headers, users_api, first_empty_user)
write_all_order_to_excel(spreadsheet, orders_gsheets, orders_gsheets_headers, order_items, id_order_dict, users_api,
                         first_empty_order)

update_prodsheets(sheets, worksheets)
