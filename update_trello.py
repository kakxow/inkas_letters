import os
import sys

import dotenv
import requests
import urllib3


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
dotenv.load_dotenv()


TRELLO_KEY = os.getenv('TRELLO_KEY')
TRELLO_TOKEN = os.getenv('TRELLO_TOKEN')
TRELLO_API = 'https://api.trello.com/1'

trello_map = {
    'Росинкас г.Белгород': '5d121cb7f63207556da89d02',
    'Росинкас г.Брянск': '5d395ca7397ac736a8131882',
    'Росинкас г.Владимир': '5d121f048ada417f6f9f24cc',
    'Росинкас Воронежское ОУИ': '5d121f274423841942135519',
    'Росинкас г.Иваново': '5d121f47f9391b32cecb6e38',
    'Росинкас г.Калуга': '5d5fd424e6010b5522856b29',
    'Костромское ОУИ': '5ccfce6a065a490984904bf6',
    'Росинкас г. Курск': '5cda611210679079c92d6b03',
    'Росинкас г.Липецк': '5cd96d851d0bd421ee724324',
    'Альфа-Банк АО': '5d122049879f306e7c8ff43b',
    'ТКБ ПАО': '5d12205bda8d693c33e64e96',
    'Росинкас г.Орел': '5d122084b5bbbb3e69b1a511',
    'Сбербанк': '5d122098461e4f109b314817',
    'Смоленское ОУИ': '5d1220adce053d497696ff6b',
    'Росинкас г.Тамбов': '5d1220bc1f5d4060eb61d491',
    'Росинкас г.Тверь': '5cdbbc52560b4c2e40f430da',
    'Росинкас г.Тула': '5d1221434ef4594ebb19e04e',
    'Росинкас г.Ярославль': '5d12215af63207556daa134e'
}


def main(ka_name: str, txt: str):
    querystring_checklists = {
        'checkItems': 'all',
        'checkItem_fields': 'none',
        'filter': 'all',
        'fields': 'id,name',
        'key': TRELLO_KEY,
        'token': TRELLO_TOKEN
    }
    querystring_checkitem = {
        'name': txt,
        'pos': 'bottom',
        'checked': 'false',
        'key': TRELLO_KEY,
        'token': TRELLO_TOKEN
    }

    url_checklists = f'{TRELLO_API}/cards/{trello_map[ka_name]}/checklists'
    response = requests.request('GET',
                                url_checklists,
                                params=querystring_checklists,
                                verify=False)
    j = response.json()
    key = 'ЕС' if ' ЕС ' in txt else 'СЛ'
    checklist = next(filter(lambda x: x['name'] == key, j))

    url_checkitem = f'{TRELLO_API}/checklists/{checklist["id"]}/checkItems'
    response_post = requests.request('POST',
                                     url_checkitem,
                                     params=querystring_checkitem,
                                     verify=False)


if __name__ == '__main__':
    ka, txt = sys.argv[1:]
    main(ka, txt)
