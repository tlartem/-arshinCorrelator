import tkinter as tk
from tkinter import filedialog
import os
import classes as cl

ARSHIN_PATH = None
JOURNAL_PATH = None
arshin_list = []
manometer_list = []



def choose_file(path_var, file_label):
    path_var = filedialog.askopenfilename(title="Выберите файл",
                                          filetypes=[("Excel files", "*.xlsx")])
    if path_var:
        # Ограничим длину отображаемого имени файла
        display_name = os.path.basename(path_var)
        if len(display_name) > 30:
            display_name = display_name[:27] + '...'
        file_label.config(text=f"Файл: {display_name}", foreground='green')


def choose_arshin():
    choose_file(ARSHIN_PATH, arshin_label)


def choose_journal():
    choose_file(JOURNAL_PATH, journal_label)


font_default = ("Tahoma", 10)

win = tk.Tk()
win.title('ArshinCorrelator')
win.geometry('400x500')
win.resizable(False, False)
win.option_add("*Font", font_default)

def do():
    xl = cl.Excel(ARSHIN_PATH, JOURNAL_PATH,
                  4,  'E', 'G', 'K',
                  )

main_frame = tk.Frame(win)
main_frame.pack(pady=20)

arshin_btn = tk.Button(main_frame, text="Файл Аршина", command=choose_arshin)
arshin_btn.grid(row=1, column=0, pady=10, columnspan=2)

arshin_label = tk.Label(main_frame, text="Файл аршина не выбран", foreground='red')
arshin_label.grid(row=0, column=0, pady=5, columnspan=2)

journal_btn = tk.Button(main_frame, text="Файл журнала", command=choose_journal)
journal_btn.grid(row=3, column=0, pady=10, columnspan=2)

journal_label = tk.Label(main_frame, text="Файл журнала не выбран", foreground='red')
journal_label.grid(row=2, column=0, pady=5, columnspan=2)

first_row_lb = tk.Label(main_frame, text="Первая строка(аршин/журнал):")
first_row_lb.grid(row=4, column=0, pady=5, columnspan=1)

arshin_first_row_entry = tk.Entry(main_frame, width=5, justify='center')
arshin_first_row_entry.grid(row=4, column=1, pady=5, columnspan=1)

journal_first_row_entry = tk.Entry(main_frame, width=5, justify='center')
journal_first_row_entry.grid(row=4, column=2, pady=5, columnspan=1)

# Копии для arshin\journal_gos_letter
gos_letter_lb = tk.Label(main_frame, text="Буква ГРСИ(аршин/журнал):")
gos_letter_lb.grid(row=5, column=0, pady=5, columnspan=1)

arshin_journal_gos_letter_entry = tk.Entry(main_frame, width=5, justify='center')
arshin_journal_gos_letter_entry.grid(row=5, column=1, pady=5, columnspan=1)

journal_gos_letter_entry = tk.Entry(main_frame, width=5, justify='center')
journal_gos_letter_entry.grid(row=5, column=2, pady=5, columnspan=1)

# Копии для number_letter
number_letter_lb = tk.Label(main_frame, text="Буква номера(аршин/журнал):")
number_letter_lb.grid(row=6, column=0, pady=5, columnspan=1)

arshin_number_letter_entry = tk.Entry(main_frame, width=5, justify='center')
arshin_number_letter_entry.grid(row=6, column=1, pady=5, columnspan=1)

journal_number_letter_entry = tk.Entry(main_frame, width=5, justify='center')
journal_number_letter_entry.grid(row=6, column=2, pady=5, columnspan=1)

# Копии для doc_letter
doc_letter_lb = tk.Label(main_frame, text="Буква документа(аршин/журнал):")
doc_letter_lb.grid(row=7, column=0, pady=5, columnspan=1)

arshin_doc_letter_entry = tk.Entry(main_frame, width=5, justify='center')
arshin_doc_letter_entry.grid(row=7, column=1, pady=5, columnspan=1)

journal_doc_letter_entry = tk.Entry(main_frame, width=5, justify='center')
journal_doc_letter_entry.grid(row=7, column=2, pady=5, columnspan=1)

do_btn = tk.Button(main_frame, text="Соотнести", command=do)
do_btn.grid(row=8, column=0, pady=5, columnspan=2)


win.mainloop()
