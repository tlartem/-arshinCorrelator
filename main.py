import tkinter as tk
from tkinter import filedialog
import os
import classes as cl
from gui_logger import initialize_console_logger

ARSHIN_PATH = None
JOURNAL_PATH = None
arshin_list = []
manometer_list = []


def choose_arshin():
    global ARSHIN_PATH
    ARSHIN_PATH = filedialog.askopenfilename(title="Выберите файл",
                                             filetypes=[("Excel files", "*.xlsx")])
    if ARSHIN_PATH:
        # Ограничим длину отображаемого имени файла
        display_name = os.path.basename(ARSHIN_PATH)
        if len(display_name) > 30:
            display_name = display_name[:27] + '...'
        arshin_label.config(text=f"Файл: {display_name}", foreground='green')


def choose_journal():
    global JOURNAL_PATH
    JOURNAL_PATH = filedialog.askopenfilename(title="Выберите файл",
                                              filetypes=[("Excel files", "*.xlsx")])
    if JOURNAL_PATH:
        # Ограничим длину отображаемого имени файла
        display_name = os.path.basename(JOURNAL_PATH)
        if len(display_name) > 30:
            display_name = display_name[:27] + '...'
        journal_label.config(text=f"Файл: {display_name}", foreground='green')


font_default = ("Tahoma", 10)

win = tk.Tk()
win.title('ArshinCorrelator')
win.geometry('400x500')
win.resizable(False, False)
win.option_add("*Font", font_default)

# Добавляем меню
win.menu = tk.Menu(win)
win.config(menu=win.menu)

win.options_menu = tk.Menu(win.menu, tearoff=0)
win.menu.add_cascade(label="Опции", menu=win.options_menu)


def show_console(*args):
    win.console_logger.show_console()


# Инициализация консольного логгера
win.console_logger = initialize_console_logger(win, show_console)
win.options_menu.add_command(label='Показать лог', command=win.console_logger.show_console)


def do():
    xl = cl.Excel(ARSHIN_PATH, JOURNAL_PATH,
                  int(arshin_first_row_entry.get()),
                  arshin_gos_letter_entry.get(),
                  arshin_number_letter_entry.get(),
                  arshin_doc_letter_entry.get(),
                  int(journal_first_row_entry.get()),
                  journal_gos_letter_entry.get(),
                  journal_number_letter_entry.get(),
                  journal_doc_letter_entry.get(),
                  True
                  )
    xl.parse_arshin()
    xl.parse_journal()
    xl.associate()
    xl.fill_doc()


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

arshin_gos_letter_entry = tk.Entry(main_frame, width=5, justify='center')
arshin_gos_letter_entry.grid(row=5, column=1, pady=5, columnspan=1)

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
