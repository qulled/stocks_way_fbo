import json
import datetime as dt
import requests
import os


def get_stocks(name, token, date_from):
    url = 'https://suppliers-stats.wildberries.ru/api/v1/supplier/stocks'
    params = {
        'key': token,
        'dateFrom': date_from,
    }
    response = requests.get(url, params=params)
    with open(f'reports/{name} {date_from}.json', 'w') as f:
        json.dump(response.json(), f, indent=2, ensure_ascii=False)
    return f'reports/{name} {date_from}.json'



if __name__ == '__main__':
    date_from = dt.datetime.date(dt.datetime.now())
    cred_file = os.path.join('credentials.json')
    with open(cred_file, 'r', encoding="utf-8") as f:
        cred = json.load(f)
    for i in cred:
        if i != 'Савельева':
            token = cred[i].get('fbo_token')
            name = i
            get_stocks(name, token, date_from)
    json_1,json_2,json_3 = f'reports/Белотелов {date_from}.json',f'reports/Орлова {date_from}.json',f'reports/Кулик {date_from}.json'
    with open(json_1,'r') as file:
        data_1 = json.load(file)
    with open(json_2,'r') as file:
        data_2 = json.load(file)
    with open(json_3,'r') as file:
        data_3 = json.load(file)
    new_data = []
    for item in data_1:
        new_data.append(item)
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), json_1)
    os.remove(path)
    for item in data_2:
        new_data.append(item)
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), json_2)
    os.remove(path)
    for item in data_3:
        new_data.append(item)
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), json_3)
    os.remove(path)
    with open(f'reports/stocks {date_from}.json','w',encoding='utf-8') as file:
        json.dump(new_data,file)

