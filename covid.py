"""
Create letters not from e-mail, but from pre-made form.
"""
from datetime import datetime
from types import MethodType

from macro_official_letter import Letter, PointOfSale


def get_data_from_input_mass2(self):
    # self.keyword = 'temp'
    self.event_date = datetime.now().date()
    print('Введите коды сап')
    while True:
        inp = input().strip('\n')
        code = inp[:4]
        if not code:
            break
        new_object = PointOfSale()
        new_object.object_SAP_code = code
        if self.keyword == 'change':
            new_object.operation_mode_new = inp[5:]
        if self.keyword == 'closing':
            new_object.successor_full_name = ''
            new_object.successor_name = ''
            date = inp[5:]
            if date:
                self.event_date = datetime.strptime(date, '%d.%m.%Y').date()
            break
        self.objects.append(new_object)


def main_mass():
    keywords = ['temp', 'reopen', 'change', 'closing']
    while True:
        letter = Letter()
        letter.objects = []
        keyword = input(', '.join(keywords) + '?').lower()
        if keyword not in keywords():
            print('ошибка ввода')
            continue
        letter.keyword = keyword
        letter._get_data_from_mail = MethodType(get_data_from_input_mass2, letter)
        letter.run()


if __name__ == '__main__':
    main_mass()
