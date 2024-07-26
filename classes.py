import random

import pythoncom
import win32com.client as win32

arshin_list = []
manometer_list = []


class Cell:
    def __init__(self, cell: str):
        cell = cell.strip()
        self.letter = cell[0]
        self.number = cell[1:]

    def __str__(self):
        return f'{self.letter}{self.number}'


def clearnumber(value):
    try:
        return str(int(value))
    except Exception as e:
        pass
    return value

class ArshinEntry:
    def __init__(self, gos, number, doc):
        self.gos = gos
        self.number = number
        self.doc = doc


class Manometer:
    def __init__(self, gos, number, row):
        self.gos = gos
        self.number = number
        self.row = row
        self.doc = ''


class Excel:
    def __init__(self, arshin_path, journal_path,
                 arshin_first_row, arshin_gos_letter, arshin_number_letter, arshin_doc_letter,
                 journal_first_row, journal_gos_letter, journal_number_letter, journal_doc_letter,
                 xl_visible=False):
        self.arshin_list = []
        self.manometer_list = []
        pythoncom.CoInitialize()
        self.counter = 0

        xlApp = win32.Dispatch('Excel.Application')
        xlApp.Visible = xl_visible

        Arshin = xlApp.Workbooks.Open(arshin_path)
        Journal = xlApp.Workbooks.Open(journal_path)

        self.arshin_com = Arshin.ActiveSheet
        self.journal_com = Journal.ActiveSheet

        self.arshin_first_row = arshin_first_row
        self.arshin_gos_letter = arshin_gos_letter
        self.arshin_number_letter = arshin_number_letter
        self.arshin_doc_letter = arshin_doc_letter

        self.journal_first_row = journal_first_row
        self.journal_gos_letter = journal_gos_letter
        self.journal_number_letter = journal_number_letter
        self.journal_doc_letter = journal_doc_letter

        # self.xlApp = xlApp
        # self.wb = Protocol

    def parse_arshin(self):
        row = self.arshin_first_row
        try:
            while row < 500:
                self.arshin_list.append(ArshinEntry(
                    self.arshin_com.Range(f'{self.arshin_gos_letter}{row}').value,
                    clearnumber(self.arshin_com.Range(f'{self.arshin_number_letter}{row}').value),
                    self.arshin_com.Range(f'{self.arshin_doc_letter}{row}').value,
                ))
                row += 1
        except Exception as e:
            print(e)

    def parse_journal(self):
        row = self.journal_first_row
        try:
            while row < 2000:
                self.manometer_list.append(Manometer(
                    self.journal_com.Range(f'{self.journal_gos_letter}{row}').value,
                    clearnumber(self.journal_com.Range(f'{self.journal_number_letter}{row}').value),
                    row
                ))
                row += 1
        except Exception as e:
            print(e)

    def associate(self):
        arshin_dict = {}

        # Создаем словарь, где ключами являются tuple (gos, number), а значениями - списки doc
        for arshin_entry in self.arshin_list:
            key = (arshin_entry.gos, arshin_entry.number)
            if key not in arshin_dict:
                arshin_dict[key] = []
            arshin_dict[key].append(arshin_entry.doc)

        # Присваиваем каждому манометру doc из списка doc для соответствующего ключа (gos, number)
        for manometer in self.manometer_list:
            key = (manometer.gos, manometer.number)

            if key in arshin_dict and arshin_dict[key]:
                manometer.doc = arshin_dict[key].pop(0)  # Берем первый доступный doc и удаляем его из списка
                self.counter += 1

        # Для вывода не найденных записей в журнале, если нужно
        for key, docs in arshin_dict.items():
            if docs:
                print(f"Не найден в журнале: {key[1]}")
        return self.counter

    def fill_doc(self):
        for manometer in self.manometer_list:
            if manometer.doc != '' and manometer.doc != None:
                self.journal_com.Range(f'{self.journal_doc_letter}{manometer.row}').value = f'{manometer.doc}'
