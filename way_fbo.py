import requests
from dotenv import load_dotenv
from googleapiclient import discovery
from google.oauth2 import service_account
from googleapiclient.discovery import build
import logging
import os
import datetime as dt
import json

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'credentials_service.json'
credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE)
service = discovery.build('sheets', 'v4', credentials=credentials)
START_POSITION_FOR_PLACE = 0

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')

if os.path.exists(dotenv_path):
    load_dotenv(dotenv_path)
load_dotenv('.env ')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')


def list_barcode(json_file):
    list_barcode = []
    with open(json_file) as f:
        templates = json.load(f)
    try:
        count =1
        for card in templates:
            if card['barcode'] not in list_barcode:
                if card['barcode'] == '':
                    card['barcode'] = f'Нет баркода{count}'.upper()
                    count+=1
                list_barcode.append(card['barcode'])
    except:
        pass
    return list_barcode


def dicts_info(json_file, list_barcode):
    dict_brand = {}
    dict_subject = {}
    dict_article = {}
    dict_size = {}
    dict_price = {}
    with open(json_file) as f:
        templates = json.load(f)
    try:
        count = 1
        for card in templates:
            if card['barcode'] == '':
                card['barcode'] = f'Нет баркода{count}'.upper()
                count += 1
            if card['barcode'] in list_barcode:
                dict_brand[card['barcode']] = card['brand']
                dict_subject[card['barcode']] = card['subject']
                dict_article[card['barcode']] = card['supplierArticle']
                dict_size[card['barcode']] = card['techSize']
                dict_price[card['barcode']] = float(card['Price']) * ((100 - int(card['Discount'])) / 100)


    except:
        pass
    return dict_brand, dict_subject, dict_article, dict_size, dict_price


def dicts_stocks(json_file):
    dict_podolsk = {}
    dict_kazan = {}
    dict_electrostal = {}
    dict_krasnodar = {}
    dict_ekb = {}
    dict_spb = {}
    dict_novosibirsk = {}
    dict_habarovsk = {}
    dict_nursultan = {}
    with open(json_file) as f:
        templates = json.load(f)
    try:
        count = 1
        for card in templates:
            if card['barcode'] == '':
                card['barcode'] = f'Нет баркода{count}'.upper()
                count += 1
            else:
                card['barcode'] = card['barcode']
            if card['warehouseName'] == 'Казань':
                dict_kazan[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Электросталь':
                dict_electrostal[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Краснодар' or card['warehouseName'] == 'Краснодар 2':
                dict_krasnodar[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Екатеринбург':
                dict_ekb[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Санкт-Петербург':
                dict_spb[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Новосибирск':
                dict_novosibirsk[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Хабаровск':
                dict_habarovsk[card['barcode']] = card['quantity']
            if card['warehouseName'] == 'Нур-Султан':
                dict_nursultan[card['barcode']] = card['quantity']
            else:
                dict_podolsk[card['barcode']] = card['quantity']
    except:
        pass
    return dict_podolsk, dict_kazan, dict_electrostal, dict_krasnodar, dict_ekb, dict_spb, dict_novosibirsk, dict_habarovsk, dict_nursultan


def dict_in_way_to_client(json_file):
    dict_in_way_to_client = {}
    with open(json_file) as f:
        templates = json.load(f)
    try:
        count = 1
        for card in templates:
            if card['barcode'] == '':
                card['barcode'] = f'Нет баркода{count}'.upper()
                count += 1
            else:
                card['barcode'] = card['barcode']
            if card['inWayToClient']!=0:
                dict_in_way_to_client[card['barcode']]=card['inWayToClient']
    except Exception as e:
        print(e)
    return dict_in_way_to_client


def dict_in_way_from_client(json_file):
    dict_in_way_from_client = {}
    with open(json_file) as f:
        templates = json.load(f)
    try:
        count = 1
        for card in templates:
            if card['barcode'] == '':
                card['barcode'] = f'Нет баркода{count}'.upper()
                count += 1
            else:
                card['barcode'] = card['barcode']
            if card['inWayFromClient']!=0:
                dict_in_way_from_client[card['barcode']]=card['inWayFromClient']
    except Exception as e:
        print(e)
    return dict_in_way_from_client


def convert_to_column_letter(column_number):
    column_letter = ''
    while column_number != 0:
        c = ((column_number - 1) % 26)
        column_letter = chr(c + 65) + column_letter
        column_number = (column_number - c) // 26
    return column_letter


def update_table_barcode(table_id, list_barcode):
    position_for_place = START_POSITION_FOR_PLACE

    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                index = 0
                if len(list_barcode) != 0:
                    value = list_barcode[index]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 4)}{i}',
                         'values': [[f'{value}']]}]
                    list_barcode.pop(index)
            except Exception as e:
                print(e)
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_table_brand(table_id, dict_brand):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_brand:
                    value = dict_brand[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 1)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_table_subject(table_id, dict_subject):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_subject:
                    value = dict_subject[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 2)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_table_article(table_id, dict_article):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_article:
                    value = dict_article[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 3)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_table_size(table_id, dict_size):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_size:
                    value = dict_size[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 5)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_table_price(table_id, dict_price):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_price:
                    value = str(dict_price[barcode]).replace('.',',')
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 6)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_to_client(table_id, dict_way):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_way:
                    value = dict_way[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 7)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


def update_from_client(table_id, dict_way):
    position_for_place = START_POSITION_FOR_PLACE
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=table_id,
                                range=range_name, majorDimension='ROWS').execute()
    values = result.get('values', [])
    i = 2
    body_data = []
    if not values:
        logging.info('No data found.')
    else:
        for row in values[1:]:
            try:
                barcode = row[3].strip().upper()
                if barcode in dict_way:
                    value = dict_way[barcode]
                    body_data += [
                        {'range': f'{range_name}!{convert_to_column_letter(position_for_place + 8)}{i}',
                         'values': [[f'{value}']]}]
            except:
                pass
            finally:
                i += 1
                body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': body_data}
    sheet.values().batchUpdate(spreadsheetId=table_id, body=body).execute()


if __name__ == '__main__':
    date_from = dt.datetime.date(dt.datetime.now())
    range_name = 'В пути'
    table_id = SPREADSHEET_ID
    json_file = f'reports/stocks {date_from}.json'
    dict_brand, dict_subject, dict_article, dict_size, dict_price = dicts_info(json_file,
                                                                               list_barcode(json_file))

    update_table_barcode(table_id, list_barcode(json_file))
    update_table_brand(table_id, dict_brand)
    update_table_subject(table_id,dict_subject)
    update_table_article(table_id,dict_article)
    update_table_size(table_id,dict_size)
    update_table_price(table_id, dict_price)
    update_to_client(table_id,dict_in_way_to_client(json_file))
    update_from_client(table_id,dict_in_way_from_client(json_file))

