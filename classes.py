import pythoncom
import pywin32.client as win32

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


class Excel:
    def __init__(self, arshin_path, journal_path,
                 arshin_first_row, arshin_gos_letter, arshin_number_letter, arshin_doc_letter,
                 journal_first_row, journal_gos_letter, journal_number_letter, journal_doc_letter,
                 xl_visible=False):
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
                arshin_list.append(ArshinEntry(
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
            while row < 200:
                arshin_list.append(Manometer(
                    self.arshin_com.Range(f'{self.journal_gos_letter}{row}').value,
                    self.arshin_com.Range(f'{self.journal_number_letter}{row}').value,
                    row
                ))
                row += 1
        except Exception as e:
            print(e)

    @staticmethod
    def associate():
        for arshin_entry in arshin_list:
            # Ищем соответствующий элемент в списке manometer_list
            for manometer in manometer_list:
                if manometer.gos == arshin_entry.gos and manometer.number == arshin_entry.number:
                    # Переносим значение doc
                    manometer.doc = arshin_entry.doc

    def fill_doc(self):
        for manometer in manometer_list:
            self.journal_com.Range(f'{self.journal_doc_letter}{manometer.row}').value = f'{manometer.doc}'
