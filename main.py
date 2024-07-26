import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import json
import classes as cl
from gui_logger import initialize_console_logger
import sv_ttk

arshin_list = []
manometer_list = []
CONFIG_FILE = 'config.json'


def choose_arshin():
    global ARSHIN_PATH
    ARSHIN_PATH = filedialog.askopenfilename(title="Выберите файл",
                                             filetypes=[("Excel files", "*.xlsx")])
    if ARSHIN_PATH:
        display_name = os.path.basename(ARSHIN_PATH)
        arshin_label.config(text=f"Файл:\n{display_name}", foreground='green')


def choose_journal():
    global JOURNAL_PATH
    JOURNAL_PATH = filedialog.askopenfilename(title="Выберите файл",
                                              filetypes=[("Excel files", "*.xlsx")])
    if JOURNAL_PATH:
        display_name = os.path.basename(JOURNAL_PATH)
        journal_label.config(text=f"Файл:\n{display_name}", foreground='green')


def save_config():
    config = {
        'arshin_first_row': arshin_first_row_entry.get(),
        'journal_first_row': journal_first_row_entry.get(),
        'arshin_gos_letter': arshin_gos_letter_entry.get(),
        'journal_gos_letter': journal_gos_letter_entry.get(),
        'arshin_number_letter': arshin_number_letter_entry.get(),
        'journal_number_letter': journal_number_letter_entry.get(),
        'arshin_doc_letter': arshin_doc_letter_entry.get(),
        'journal_doc_letter': journal_doc_letter_entry.get(),
        'journal_additional_letter': journal_additional_letter_entry.get(),
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            arshin_first_row_entry.insert(0, config.get('arshin_first_row', ''))
            journal_first_row_entry.insert(0, config.get('journal_first_row', ''))
            arshin_gos_letter_entry.insert(0, config.get('arshin_gos_letter', ''))
            journal_gos_letter_entry.insert(0, config.get('journal_gos_letter', ''))
            arshin_number_letter_entry.insert(0, config.get('arshin_number_letter', ''))
            journal_number_letter_entry.insert(0, config.get('journal_number_letter', ''))
            arshin_doc_letter_entry.insert(0, config.get('arshin_doc_letter', ''))
            journal_doc_letter_entry.insert(0, config.get('journal_doc_letter', ''))
            journal_additional_letter_entry.insert(0, config.get('journal_additional_letter', ''))


def on_closing():
    save_config()
    win.destroy()


def show_console(*args):
    win.console_logger.show_console()


def do():
    win.console_logger.clear_textbox()
    xl = cl.Excel(ARSHIN_PATH, JOURNAL_PATH,
                  int(arshin_first_row_entry.get()),
                  arshin_gos_letter_entry.get(),
                  arshin_number_letter_entry.get(),
                  arshin_doc_letter_entry.get(),
                  int(journal_first_row_entry.get()),
                  journal_gos_letter_entry.get(),
                  journal_number_letter_entry.get(),
                  journal_doc_letter_entry.get(),
                  journal_additional_letter_entry.get(),
                  True
                  )
    xl.parse_arshin()
    xl.parse_journal()
    counter = xl.associate()
    xl.fill_doc()
    messagebox.showinfo(f'Соотнесено {counter}шт.')


win = tk.Tk()
win.title('ArshinCorrelator')
win.geometry('400x550')
win.resizable(False, False)
win.option_add("*Font", ("Tahoma", 12))
sv_ttk.set_theme("light")
win.iconbitmap('_internal\\icon.ico')

# Инициализация консольного логгера
win.console_logger = initialize_console_logger(win, show_console)

# Меню
win.menu = tk.Menu(win)
win.config(menu=win.menu)
win.options_menu = tk.Menu(win.menu, tearoff=0)
win.menu.add_cascade(label="Опции", menu=win.options_menu)
win.options_menu.add_command(label='Показать лог', command=win.console_logger.show_console)

main_frame = ttk.Frame(win)
main_frame.pack(pady=20)

arshin_btn = ttk.Button(main_frame, text="Файл Аршина", command=choose_arshin)
arshin_btn.grid(row=1, column=0, pady=10, columnspan=3)

arshin_label = ttk.Label(main_frame, text="Файл аршина не выбран", foreground='red', wraplength=250)
arshin_label.grid(row=0, column=0, pady=5, columnspan=3)

journal_btn = ttk.Button(main_frame, text="Файл Журнала", command=choose_journal)
journal_btn.grid(row=3, column=0, pady=10, columnspan=3)

journal_label = ttk.Label(main_frame, text="Файл журнала не выбран", foreground='red', wraplength=250)
journal_label.grid(row=2, column=0, pady=5, columnspan=3)

first_row_lb = ttk.Label(main_frame, text="Первая строка (аршин/журнал):")
first_row_lb.grid(row=4, column=0, pady=5, columnspan=1)

arshin_first_row_entry = ttk.Entry(main_frame, width=5, justify='center')
arshin_first_row_entry.grid(row=4, column=1, pady=5, columnspan=1)

journal_first_row_entry = ttk.Entry(main_frame, width=5, justify='center')
journal_first_row_entry.grid(row=4, column=2, pady=5, columnspan=1)

# Копии для arshin\journal_gos_letter
gos_letter_lb = ttk.Label(main_frame, text="Буква ГРСИ (аршин/журнал):")
gos_letter_lb.grid(row=5, column=0, pady=5, columnspan=1)

arshin_gos_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
arshin_gos_letter_entry.grid(row=5, column=1, pady=5, columnspan=1)

journal_gos_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
journal_gos_letter_entry.grid(row=5, column=2, pady=5, columnspan=1)

# Копии для number_letter
number_letter_lb = ttk.Label(main_frame, text="Буква номера (аршин/журнал):")
number_letter_lb.grid(row=6, column=0, pady=5, columnspan=1)

arshin_number_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
arshin_number_letter_entry.grid(row=6, column=1, pady=5, columnspan=1)

journal_number_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
journal_number_letter_entry.grid(row=6, column=2, pady=5, columnspan=1)

# Копии для doc_letter
doc_letter_lb = ttk.Label(main_frame, text="Буква документа (аршин/журнал):")
doc_letter_lb.grid(row=7, column=0, pady=5, columnspan=1)

arshin_doc_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
arshin_doc_letter_entry.grid(row=7, column=1, pady=5, columnspan=1)

journal_doc_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
journal_doc_letter_entry.grid(row=7, column=2, pady=5, columnspan=1)

# Дополнительное поле для journal
additional_letter_lb = ttk.Label(main_frame, text="Доп. буква (журнал):")
additional_letter_lb.grid(row=8, column=0, pady=5, columnspan=1)

journal_additional_letter_entry = ttk.Entry(main_frame, width=5, justify='center')
journal_additional_letter_entry.grid(row=8, column=2, pady=5, columnspan=1)

do_btn = ttk.Button(main_frame, text="Соотнести", command=do)
do_btn.grid(row=9, column=0, pady=5, columnspan=3)

load_config()

win.protocol("WM_DELETE_WINDOW", on_closing)
win.mainloop()