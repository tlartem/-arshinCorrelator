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
            while row < 200:
                self.arshin_list.append(ArshinEntry(
                    self.arshin_com.Range(f'{self.arshin_gos_letter}{row}').value,
                    self.arshin_com.Range(f'{self.arshin_number_letter}{row}').value,
                    self.arshin_com.Range(f'{self.arshin_doc_letter}{row}').value,
                ))
                row += 1
        except Exception as e:
            print(e)

    def parse_journal(self):
        row = self.journal_first_row
        try:
            while row < 1000:
                self.manometer_list.append(Manometer(
                    self.journal_com.Range(f'{self.journal_gos_letter}{row}').value,
                    self.journal_com.Range(f'{self.journal_number_letter}{row}').value,
                    row
                ))
                row += 1
        except Exception as e:
            print(e)

    def associate(self):
        arshin_list_copy = self.arshin_list[:]

        for arshin_entry in arshin_list_copy:
            for manometer in self.manometer_list:
                manometer_number = manometer.number
                arshin_number = arshin_entry.number

                # Преобразуем номера в строки без пробелов
                manometer_number_str = str(manometer_number).strip()
                arshin_number_str = str(arshin_number).strip()

                # Если оба номера числовые, приводим их к int и сравниваем
                try:
                    if float(manometer_number) == float(arshin_number):
                        manometer_number_str = str(int(float(manometer_number)))
                        arshin_number_str = str(int(float(arshin_number)))
                except ValueError:
                    # Один из номеров не числовой, продолжаем обычное сравнение строк
                    pass

                if manometer.gos == arshin_entry.gos and manometer_number_str == arshin_number_str:
                    manometer.doc = arshin_entry.doc
                    self.arshin_list.remove(arshin_entry)
                    break

        if self.arshin_list:
            for entry in self.arshin_list:
                print(f"Не найден в журнале: {entry.number}")

    def fill_doc(self):
        for manometer in self.manometer_list:
            if manometer.doc != '' and manometer.doc != None:
                self.journal_com.Range(f'{self.journal_doc_letter}{manometer.row}').value = f'{manometer.doc}'
