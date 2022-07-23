import json
from pathlib import Path
import pandas as pd
import datetime as dt
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
cred_file = os.path.join(BASE_DIR, 'credentials.json')
date_from = dt.datetime.date(dt.datetime.now())
day = dt.datetime.now().strftime("%d")
month = dt.datetime.now().strftime("%m")
year = dt.datetime.now().year
path = Path(r'C:\Users\ikaty\PycharmProjects\stocks_way_FBO\excel_docs')

with open(cred_file, 'r', encoding="utf-8") as f:
    cred = json.load(f)
    for name in cred:
        if name != 'Савельева' and name !='Кулик':
            df = pd.read_excel(f'{path}/{name}-{date_from}.xlsx')
            df.insert(0,'Юр лицо',f'ИП {name}')
            df.to_excel(f'{path}/{name}-{date_from}.xlsx')

df = pd.concat([pd.read_excel(f) for f in path.glob("*.xlsx")],
               ignore_index=True)
df.to_excel(rf'C:\Users\ikaty\PycharmProjects\stocks_way_FBO\final_excel\ALL-{date_from}.xlsx')

with open(cred_file, 'r', encoding="utf-8") as f:
    cred = json.load(f)
    for name in cred:
        if name != 'Савельева':
            file = os.path.join(f'{path}/{name}-{date_from}.xlsx')
            os.remove(file)