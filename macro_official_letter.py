import re
import csv
import os
import locale
import json
from datetime import datetime, timedelta
import datetime as dt
from typing import Dict, Tuple, List, Union
from contextlib import contextmanager

import win32com.client as win32
import win32clipboard as clipboard
from bs4 import BeautifulSoup
from jinja2 import Template

# import sberinkas
import terminals
import update_trello
from macro_official_letter_cfg import *


class PointOfSale:
    object_code: str
    object_SAP_code: str
    object_name: str
    object_address: str
    object_fed_subj: str
    object_head: str
    object_head_phone: str
    lat: str
    lon: str
    operation_mode_old: str
    operation_mode_new: str
    counterparty_name: str
    contract: str
    ter_dir_name: str
    ter_dir_email: str
    successor_name: str
    successor_full_name: str
    successor_email: str = ''
    terminals: List[str]


class Letter():
    event_date: dt.date
    signatory: str
    letter_name: str
    letter_text: str
    attachment: str
    keyword: str
    path: str
    to: Union[str, List[str]]
    template: str
    objects: List[PointOfSale] = []

    def __init__(self):
        """
        Opens MS Outlook on init, sets locale, determines keyword and mail_text.
        """
        self.outlook = win32.Dispatch('outlook.application')
        locale.setlocale(locale.LC_ALL, '')

    def run(self):
        self.get_all_data()
        self._prepare_letter()
        self._create_letter(TEMP_FILE, TEMPLATE_PATH, print_flag=False)
        self.send_all()
        self.trello()

    def get_all_data(self):
        with open(SIGNATORY_FILE) as f:
            self.signatory = f.readline()
        self._get_data_from_mail()
        self._contract_from_file()
        self._info_from_nsi()
        self._check_address()
        self._letter_info()
        if self.keyword == 'closing':
            self._get_terminals()

    def trello(self):
        if self.keyword in ('closing', 'change'):
            for object_ in self.objects:
                update_trello.main(object_.counterparty_name, self.letter_text)

    def _get_data_from_mail(self) -> None:
        """
        Gets valuable info from mail_text.
        """
        current_email = self.outlook.ActiveInspector()
        cap = current_email.Caption
        try:
            self.keyword = next(KEYWORDS[key] for key in KEYWORDS if key in cap)
        except StopIteration:
            raise RuntimeError('This e-mail is not standard.')

        self.mail_text = current_email.CurrentItem.HTMLBody
        # print(self.mail_text)
        if not self.mail_text:
            raise RuntimeError('Empty e-mail.')
        event_key = EVENT_KEYS.get(self.keyword, False)

        parsed_html = BeautifulSoup(self.mail_text, features="lxml")
        rows = parsed_html.find_all('tr')
        headers = parsed_html.find_all('h1')

        def table_find(text: str) -> str:
            try:
                row = next(filter(lambda x: text == x.td.string, rows))
            except StopIteration:
                row = next(filter(lambda x: text == x.td.p.text, rows))
            return row.find_all('td')[1].text

        def header_find(text: str) -> str:
            header = next(filter(lambda x: x.text.startswith(text), headers))
            return header.text

        new_object = PointOfSale()
        # Data from the table in mail body.
        new_object.object_code = table_find('Код')
        new_object.object_SAP_code = table_find('Код ТТ SAP')
        # self.object_name = table_find('Наименование')
        # self.object_address = table_find('Адрес')
        new_object.ter_dir_name = table_find('Оперативный менеджер ТТ')
        # Data from h1 tags under the table (for closing)
        if self.keyword == 'closing':
            cms = header_find('ЦМС')
            new_object.successor_full_name = cms[:cms.index(' принимает')]
            new_object.successor_name = \
                re.match(r'ЦМС \d+ (.*)', new_object.successor_full_name).group(1)
        if self.keyword == 'change':
            new_object.operation_mode_new = table_find('Режим работы')
        if event_key:
            date = re.search(
                r'\d{1,2}.\d{2}.\d{2,4}',
                header_find(event_key)
            ).group(0)
            self.event_date = \
                max(datetime.strptime(date, '%d.%m.%Y').date(), datetime.now().date())
        else:
            self.event_date = datetime.now().date()

        self.objects.append(new_object)

    def _contract_from_file(self) -> None:
        """
        Gets counterparty info from Всё про ЦР file and dict-file:
        counterparty_name
        contract
        path
        to
        template
        """
        for object_ in self.objects:
            with open(DATA_FILE) as f:
                reader = csv.reader(f, delimiter=';')
                SAP_code = object_.object_SAP_code
                for line in reader:
                    if line[0] == SAP_code:
                        object_.object_code = line[1]
                        object_.counterparty_name = line[6]
                        object_.contract = line[5]
                        break
                else:
                    raise RuntimeError('No counterparty or contract found')
            # Check for Sberbank.
            if re.search('сб', object_.counterparty_name, re.I):
                object_.counterparty_name = 'Сбербанк'
            # counterparty_info = \
            #     self.create_counterparty_dict(COUNTERPARTIES)[self.counterparty_name]
            # self.path, self.to = counterparty_info.values()
            self.path, self.to = \
                self._counterparty(COUNTERPARTIES, object_.counterparty_name)
            self.template = self.path + 'template (НЕ УДАЛЯТЬ).html'

    def _info_from_nsi(self) -> None:
        """
        Gets info from NSI file:
        object_email
        operation_mode_old
        ter_dir_email
        successor_email
        """
        column_list = [
            'object_name',
            'object_address',
            'object_head',
            'object_head_phone',
            'operation_mode_old',
            'ter_dir_email',
            'lat',
            'lon',
            'object_fed_subj',
        ]
        for object_ in self.objects:
            with open(NSI_FILE) as f:
                reader = csv.reader(f, delimiter=';')
                succ_flag = True if self.keyword == 'closing' else False
                obj_flag = True
                for line in reader:
                    if not succ_flag and not obj_flag:
                        break
                    if obj_flag:
                        if line[NSI.object_code] == object_.object_code:
                            for column_name in column_list:
                                column_no = getattr(NSI, column_name)
                                setattr(object_, column_name, line[column_no])
                            obj_flag = False
                    if succ_flag:
                        if line[NSI.successor_name] == object_.successor_name:
                            object_.successor_email = line[NSI.successor_email]
                            succ_flag = False

    def _check_address(self):
        """
        Fix address for Euroset objects.
        """
        for object_ in self.objects:
            if object_.object_name.endswith(' ЕС'):
                if object_.object_address[:6].isnumeric():
                    object_.object_address = \
                        object_.object_address[:7] + \
                        object_.object_fed_subj + ', ' + \
                        object_.object_address[7:]

    def _letter_info(self) -> None:
        """
        Gets letter texts and files from LETTER_FIXTURE:
        letter_name
        letter_text
        letter_file
        attachment
        """
        if self.keyword == 'closing':
            self.event_date += timedelta(1)
        self.letter_file, letter_name, letter_text = LETTER_FIXTURE[self.keyword]
        if len(self.objects) > 1:
            object_ = self.objects[0]
            letter_part = f" {len(self.objects)} ТТ {object_.object_fed_subj}"
        else:
            object_ = self.objects[0]
            letter_part = f" ТТ {object_.object_SAP_code} {object_.object_name}"
        today = datetime.now().strftime('%d.%m.%Y')
        self.letter_name = letter_name + letter_part + f" от {today}"
        self.letter_text = letter_text + letter_part + f" c {self.event_date.strftime('%d.%m.%Y')}"
        self.attachment = PDF_SAVE_PATH / f'{self.letter_name}.pdf'

    def _get_terminals(self):
        for object_ in self.objects:
            object_.terminals = terminals.get(object_.object_SAP_code)

    @staticmethod
    def create_counterparty_dict(file_name) -> Dict[str, str]:
        """
        Creates specific dictionary from file
        """
        dct = {}
        with open(file_name) as f:
            root_dir = f.readline().strip('\n')
            for line in f:
                key, val = line.strip('\n').split('!!!!')
                temp = val.split('==')
                d = {'path': root_dir + temp[0], 'to': temp[1:]}
                dct[key] = d
        return dct

    @staticmethod
    def _counterparty(
        file_name: str,
        cp_name: str
    ) -> Tuple[str, List[str]]:
        """
        Loads json, finds path to directory and list of e-mail addresses for
         given counterparty (cp_name).

        Parameters
        ----------
        file_name
            Name of .json file.
        cp_name
            Counterparty name.

        Returns
        -------
        Tuple[str, List[str]]
            Tuple of path to counterparty directory and e-mail addresses.
        """
        with open(file_name, encoding='utf-8') as f:
            dct = json.load(f)
        root_dir = dct.pop('root_dir')
        item = dct.pop(cp_name)
        return root_dir + item['path'], item['to']

    def _prepare_letter(self) -> None:
        """
        Prepare letter text - combine template with file, format and save as temp.
        """
        template = self.template
        letter_file = self.letter_file
        date_string = datetime.now().strftime('%d %B %Y')
        with open(template) as t, open(letter_file, encoding='utf-8') as f:
            text = t.read() + '\n' + f.read()
        letter_template = Template(text)
        template_data = {
            'date': self.event_date.strftime('%d.%m.%Y'),
            'date1': (self.event_date - timedelta(1)).strftime('%d.%m.%Y'),
            'signatory': self.signatory,
            'contract': re.sub('договор', '', self.objects[0].contract, flags=re.I),
            'objects': self.objects,
        }
        text = letter_template.render(template_data)
        replacements = {
            "tab": 'tab',
            "today": date_string,
        }
        text = text.format(**replacements).replace("&quot;", '"')
        replace_pattern = r'</body>\s+</html>\s+<!DOCTYPE html>\s+<html>\s+<body>'
        text = re.sub(replace_pattern, '', text, flags=re.I)
        with open(TEMP_FILE, 'w') as f:
            f.write(text)

    def _create_letter(
        self,
        temp_file: str,
        template_path: str,
        print_flag: bool
    ) -> None:
        """
        Creates letter in word, saves it as .docx and .pdf.
        """
        # breakpoint()
        with self.word_app() as word:
            doc = word.documents.Add()
            doc.Content.InsertFile(temp_file)
            # self.insert_file(word, doc, temp_file)
            doc.Content.Select()
            word.Selection.ParagraphFormat.TabStops.Add(
                Position=453.54,
                Alignment=2)
            word.Selection.Font.Name = FONT_NAME
            word.Selection.Font.size = FONT_SIZE
            word.Application.Run('tb')
            doc.Content.Copy()

            final_doc = word.documents.Add(template_path)
            final_doc.Content.Select()
            word.Selection.Collapse(0)
            word.Selection.Paste()
            doc.Close(SaveChanges=0)
            final_doc.SaveAs(self.path + self.letter_name + ".docx")
            # final_doc.SaveAs2(self.attachment, 17)
            final_doc.SaveAs2(self.attachment, 17)
            if print_flag:
                final_doc.PrintOut(Background=True, Copies=1)
            final_doc.Close()
        os.remove(temp_file)

    @staticmethod
    @contextmanager
    def word_app():
        """
        MS Word with context manager.
        """
        word = win32.Dispatch('Word.Application')
        try:
            yield word
        finally:
            clipboard.OpenClipboard()
            clipboard.EmptyClipboard()
            clipboard.CloseClipboard()
            word.Quit()

    @staticmethod
    def insert_file(wrd, doc, filename: str) -> None:
        """
        Insert file contents to word document.
        """
        doc.Content.Select()
        wrd.Selection.Collapse(0)
        wrd.Selection.InsertFile(filename)

    def send_all(self) -> None:
        self.send_counterparty()
        if self.keyword == 'closing':
            self.send_object()
            self.send_sber()

    def send_counterparty(self) -> None:
        """
        Generate and send e-mail to counterparty
        """
        object_ = self.objects[0]
        ticket_text = ''
        if 'сб' in object_.counterparty_name.lower() and self.keyword == 'closing':
            # order_id = sberinkas.main(
            #     object_.object_SAP_code,
            #     object_.object_address,
            #     object_.lat,
            #     object_.lon
            # )
            # ticket_text = f"<br>Номер заявки на портале инкассация - {order_id}."
            pass

        body = '<p>Добрый день!<br><br>' \
               f'Прошу принять в работу письмо на {self.letter_text}<br>' \
               f'Скан подписанного письма вышлю позднее.{ticket_text}'
        if 'сб' in object_.counterparty_name.lower():
            self.send_sber_manager_service(body)
        else:
            self.sendmail(
                self.outlook,
                self.to,
                "",
                self.letter_name,
                body,
                self.attachment,
                2
            )

    @staticmethod
    def sendmail(
        outlook,
        strTo: Union[str, List[str]],
        strCC: str,
        strSubject: str,
        str_body: str,
        strAttPathName: str,
        importance: int = 1,
        send: bool = False
    ) -> None:
        """
        importance = {'High': 2, 'Normal': 1, 'Low': 0,}
        """
        def wrap_html(txt: str) -> str:
            res = \
                '<html><head>'\
                '<style>p {font-size: 11pt; font-family: "Calibri";}</style>'\
                '</head><body>{txt}</body></html>'.replace('{txt}', txt)
            return res

        mail = outlook.Application.CreateItem(0)
        tmp = wrap_html(str_body)
        mail.Display()
        signature = mail.HTMLBody
        if isinstance(strTo, str):
            strTo = [strTo]
        for to in strTo:
            mail.Recipients.Add(to)
        mail.CC = strCC
        mail.BCC = ''
        mail.Subject = strSubject
        mail.HTMLBody = tmp + '<br>' + signature
        if strAttPathName != "":
            mail.Attachments.Add(strAttPathName)
        mail.importance = importance
        if send:
            mail.Send()

    def send_object(self):
        """
        Send an email about closing to the object.
        """
        for object_ in self.objects:
            strCC = '; '.join([object_.ter_dir_email, object_.successor_email])
            strCC += "; ekb.inkas.net@maxus.ru; schugunov@svyaznoy.ru"
            strSubject = "Инкассация и вывоз POS-терминала при закрытии ТТ"
            outMail = self.outlook.Application.CreateItemFromTemplate(
                CLOSING_MAIL_TEMPLATE
            )
            fixture = {
                'дата+1': self.event_date.strftime('%d.%m.%Y'),
                'преемник': object_.successor_full_name,
                'имяТТ': f'ЦМС {object_.object_code[-4:]} {object_.object_name}'
            }
            HTML_body_without_signature = outMail.HTMLBody
            outMail.Display()
            for k, v in fixture.items():
                HTML_body_without_signature = HTML_body_without_signature.replace('{' + k + '}', v)

            outMail.HTMLBody = HTML_body_without_signature
            outMail.To = object_.object_SAP_code
            outMail.CC = strCC
            outMail.Subject = strSubject
            outMail.importance = 2
            if datetime.now().date() + timedelta(days=1) < self.event_date:
                outMail.DeferredDeliveryTime = \
                    (self.event_date - timedelta(days=1)).strftime('%d.%m.%Y') + " 17:00"

    def send_sber(self):
        """
        E-mail to Sber for them to requesite their POS-terminal.
        """
        if len(self.objects) == 1:
            object_ = self.objects[0]
            subject = 'Закрытие ТТ и вывоз терминала ' \
                      f'{object_.object_SAP_code} {object_.object_name}'
        else:
            subject = 'Закрытие ТТ и вывоз терминалов'
        with open(SBER_TEMPLATE, encoding='utf-8') as f:
            template_text = f.read()
        template = Template(template_text)
        body = template.render(objects=self.objects, date=self.event_date.strftime("%d.%m.%Y"))
        # body = \
        #     '<p>Добрый день!</p>' \
        #     '<p>В связи с закрытием ТТ '\
        #     f'{self.object_SAP_code} {self.object_name}, ' \
        #     'прошу организовать вывоз терминала в первой половине дня '\
        #     f'{self.event_date.strftime("%d.%m.%Y")}, адрес ТТ - {self.object_address}' \
        #     '<br>Заранее спасибо!</p>'

        self.sendmail(
            self.outlook,
            ['VASaparkina@sberbank.ru', 'RAMinazhetdinov@sberbank.ru'],
            'schugunov@svyaznoy.ru',
            subject,
            body,
            '',
            1
        )

    def send_sber_manager_service(self, body: str):
        mail_subject = '7714617793/997350001 ООО "Сеть Связной" ' + self.letter_name
        self.sendmail(
            self.outlook,
            self.to,
            "",
            mail_subject,
            body,
            self.attachment,
            2
        )


if __name__ == '__main__':
    gg = Letter()
    gg.run()

    # pprint(create_counterparty_dict())

    # word = win32.Dispatch('Word.Application')
    # print(word.Application.Run('tb'))
