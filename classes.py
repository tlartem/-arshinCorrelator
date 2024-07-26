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
    def __init__(self, gos, number, doc, row_data):
        self.gos = gos
        self.number = number
        self.doc = doc
        self.row_data = row_data  # Поле для хранения всей строки


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
        self.not_found_in_journal = 0
        self.not_found_in_arshin = 0

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

    def parse_arshin(self):
        row = self.arshin_first_row
        try:
            while True:
                gos = self.arshin_com.Range(f'{self.arshin_gos_letter}{row}').value
                number = self.arshin_com.Range(f'{self.arshin_number_letter}{row}').value
                doc = self.arshin_com.Range(f'{self.arshin_doc_letter}{row}').value

                if not gos and not number and not doc:
                    print(f"Парсинг аршина остановлен на строке {row}")
                    break

                row_data = self.arshin_com.Range(f'{self.arshin_gos_letter}{row}:{self.arshin_doc_letter}{row}').Value
                self.arshin_list.append(ArshinEntry(gos, clearnumber(number), doc, row_data))
                row += 1
        except Exception as e:
            print(e)

    def parse_journal(self):
        row = self.journal_first_row
        try:
            while True:
                gos = self.journal_com.Range(f'{self.journal_gos_letter}{row}').value
                number = self.journal_com.Range(f'{self.journal_number_letter}{row}').value

                if not gos and not number:
                    print(f"Парсинг журнала остановлен на строке {row}")
                    break

                self.manometer_list.append(Manometer(gos, clearnumber(number), row))
                row += 1
        except Exception as e:
            print(e)

    def associate(self):
        arshin_dict = {}

        for arshin_entry in self.arshin_list:
            key = (arshin_entry.gos, arshin_entry.number)
            if key not in arshin_dict:
                arshin_dict[key] = []
            arshin_dict[key].append((arshin_entry.doc, arshin_entry.row_data))

        for manometer in self.manometer_list:
            key = (manometer.gos, manometer.number)

            if key in arshin_dict and arshin_dict[key]:
                manometer.doc, row_data = arshin_dict[key].pop(0)
                self.counter += 1
            else:
                self.not_found_in_arshin += 1

        for key, docs in arshin_dict.items():
            if docs:
                self.not_found_in_journal += len(docs)
                for doc, row_data in docs:
                    print(f"Не найден в журнале: {key[1]} | Строка данных: {row_data}")

        total_arshin_entries = len(self.arshin_list)
        total_journal_entries = len(self.manometer_list)
        found_in_journal = self.counter
        not_found_in_journal = self.not_found_in_journal
        not_found_in_arshin = self.not_found_in_arshin

        print(f"Общее количество записей в Аршине: {total_arshin_entries}")
        print(f"Общее количество записей в Журнале: {total_journal_entries}")
        print(f"Количество найденных совпадений: {found_in_journal}")
        print(f"Количество записей, не найденных в Журнале: {not_found_in_journal}")
        print(f"Количество записей в Журнале, не найденных в Аршине: {not_found_in_arshin}")

        return self.counter

    def fill_doc(self):
        for manometer in self.manometer_list:
            if manometer.doc != '' and manometer.doc is not None:
                self.journal_com.Range(f'{self.journal_doc_letter}{manometer.row}').value = f'{manometer.doc}'
