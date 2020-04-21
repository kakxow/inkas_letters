from datetime import datetime, timedelta
import json
import os
import sys

import dotenv
import requests


dotenv.load_dotenv()
login = os.getenv('SBER_LOGIN')
password = os.getenv('SBER_PASSWORD')
date = datetime.now() + timedelta(days=3)

url_login = 'https://sberbank-inkassacia.ru/api/user/login/'
url_add = 'https://sberbank-inkassacia.ru/api/order/contracts/add'
url_object = 'https://sberbank-inkassacia.ru/api/change/object/edit/'
url_date = 'https://sberbank-inkassacia.ru/api/change/edit'
url_company = 'https://sberbank-inkassacia.ru/api/order/about/'

login_creds = {
    "email": login,
    "password": password,
}
contract_sl = {
    "order_type": 1,
    "contracts": [
        {
            "number": "54/12/129",
            "date": "2012-12-10T00:00:00.000Z",
            "service_package": 0,
            "service": 0,
        },
        {
            "number": "54/12/130",
            "date": "2012-12-10T00:00:00.000Z",
            "service_package": 0,
            "service": 1,
        },
    ]
}
contract_es = {
    "order_type": 1,
    "contracts": [
        {
            "number": "124/16/02",
            "date": "2016-02-25T00:00:00.000Z",
            "service_package": 0,
            "service": 0,
        },
        {
            "number": "124/16/01",
            "date": "2016-02-25T00:00:00.000Z",
            "service_package": 0,
            "service": 1,
        },
    ]
}
address_json = {
        "order_id": '',
        "change_type": "1",
        "params": {
            "address": {
                "address_line": '',
                "office": '',
                "add_info": '',
                "lat": '',
                "lon": '',
            },
            "id": '',
        },
}
date_json = {
    "order_id": '',
    "change_type": "1",
    "params": '',
    "date": date.strftime('%Y-%m-%dT09:00:00.000Z'),
}
company_json = {
    "organization_name": "ООО \"Сеть Связной\"",
    "inn": "7714617793",
    "kpp": "997350001",
    "legal_address": "123007, Москва, 2-ой Хорошевский проезд, д. 9, корпус 2, этаж 5, комната 4",
    "physical_address": "115280, Москва, ул. Ленинская Слобода, д. 26с5, БЦ «Симонов Плаза»",
    "scope": '',
    "order_id": '',
}


def check_r(r) -> bool:
    body = json.loads(r.text)
    print(body)
    return body['result'] and not body['error_code'] and r.status_code == 200


def main(
    object_code: str,
    address: str,
    lat: str,
    lon: str
) -> str:
    address_json.update({
        "params": {
            "address": {
                "address_line": address,
                "office": None,
                "add_info": object_code,
                "lat": lat,
                "lon": lon,
            }
        }
    })
    with requests.Session() as s:
        # Login:
        r_login = s.post(url_login, json=login_creds)
        assert check_r(r_login)
        # Create new order and get order_id:
        r_add = s.post(url_add, json=contract_sl)
        assert r_add.status_code == 200
        body = json.loads(r_add.text)
        order_id = body['result'].get('order_id', None)
        assert order_id and not body['error_code']
        # Add date of action:
        date_json['order_id'] = order_id
        r_date = s.post(url_date, json=date_json)
        assert check_r(r_date)
        # Add an object to order.
        address_json['order_id'] = order_id
        r_object = s.post(url_object, json=address_json)
        breakpoint()
        assert check_r(r_object), r_object.text
        # Add company credentials:
        company_json['order_id'] = order_id
        r_company = s.post(url_company, json=company_json)
        assert check_r(r_company)
    return order_id


if __name__ == '__main__':
    args = sys.argv
    object_code, address, lat, lon = args[1:]
    order = main(object_code, address, lat, lon)
    print(order)
