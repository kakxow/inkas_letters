import os
from pathlib import Path
from collections import namedtuple


__all__ = ['DATA_FILE', 'TEMPLATE_PATH', 'TEMP_FILE', 'NSI_FILE', 'NSI',
           'COUNTERPARTIES', 'KEYWORDS', 'EVENT_KEYS', 'LETTER_FIXTURE',
           'FONT_NAME', 'FONT_SIZE', 'SBER_TEMPLATE', 'CLOSING_MAIL_TEMPLATE',
           'SIGNATORY_FILE', 'PDF_SAVE_PATH']


def get_nsi_file():
    a = []
    for i, entry in enumerate(os.scandir(r'C:\Max')):
        file_name = entry.name.lower()
        if file_name.endswith('.csv') and \
                file_name.startswith('нси') and \
                'общий' not in file_name:
            a.append(entry)
    file = max(a, key=lambda x: x.stat().st_ctime)

    return file.path


FONT_NAME = 'Svyaznoy Sans Light'
FONT_SIZE = 10

COUNTERPARTIES = r'C:\Max\letters\dict.json'
DATA_FILE = r'C:\Max\всё про ЦР.csv'
TEMPLATE_PATH = r"C:\Max\letters\templates\Шаблон оф.письма.dotx"
TEMP_FILE = r'c:\Max\temp_file.html'
CLOSING_MAIL_TEMPLATE = r"C:\Max\letters\templates\Шаблон закрытие ТТ.oft"
SIGNATORY_FILE = r'C:\Max\letters\signer.txt'
PDF_SAVE_PATH = Path(r'C:\Max\На подпись Кондрашовой')
SBER_TEMPLATE = r'C:\Max\letters\templates\sber_mail_template.html'


NSI_FIELDS = {
    'object_code': 0,
    'object_SAP_code': False,
    'object_name': 1,
    'object_email': 44,
    'object_address': 4,
    'object_head': 35,
    'object_head_phone': 39,
    'operation_mode_old': 34,
    'ter_dir_name': 31,
    'ter_dir_email': 32,
    'successor_name': 1,
    'successor_email': 44,
    'lat': 6,
    'lon': 7,
    'object_fed_subj': 29,
}

KEYWORDS = {
    'Возобновление': 'reopen',
    'Временное закрытие': 'temp',
    'Закрытие': 'closing',
    'Изменение режима': 'change',
}
EVENT_KEYS = {
    'closing': 'Последний рабочий день',
    'reopen': 'Открыта после',
    'temp': 'Временно закры',
}
LETTER_FIXTURE = {
    'closing': (
        r'C:\Max\letters\templates\close_letter.html',
        'Cнятие с обслуживания',
        'cнятие с обслуживания',
    ),
    'temp': (
        r'C:\Max\letters\templates\temp_letter.html',
        'Временная приостановка инкассации',
        'временную приостановку инкассации',
    ),
    'reopen': (
        r'C:\Max\letters\templates\reopen_letter.html',
        'Возобновление инкассации',
        'возобновление инкассации',
        ),
    'change': (
        r'C:\Max\letters\templates\change_letter.html',
        'Изменение режима работы',
        'изменение режима работы',
    ),
}

NSI_FILE = get_nsi_file()
nsi = namedtuple('NSI', list(NSI_FIELDS))
NSI = nsi(*NSI_FIELDS.values())
