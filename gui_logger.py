import tkinter as tk
from tkinter import scrolledtext
import sys


class TextRedirector:
    def __init__(self, text_widget, update_status_callback=None, stream_type='stdout'):
        self.text_widget = text_widget
        self.update_status_callback = update_status_callback
        self.stream_type = stream_type

    def write(self, message):
        self.text_widget.configure(state='normal')  # Временно включаем редактирование
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')  # Отключаем редактирование
        if self.update_status_callback and self.stream_type == 'stderr':
            self.update_status_callback(message)

    def flush(self):
        pass  # Этот метод нужен для совместимости с объектом sys.stdout и sys.stderr


class ConsoleLoggerApp:
    def __init__(self, root, update_status_callback=None):
        self.console_window = tk.Toplevel(root)
        self.console_window.title("Console")
        self.console_window.iconbitmap('_internal\\icon.ico')
        self.console_window.geometry("600x400")
        self.console_window.withdraw()  # Скрываем консольное окно при запуске
        self.console_window.protocol("WM_DELETE_WINDOW", self.hide_console)  # Обработка закрытия окна

        self.textbox = scrolledtext.ScrolledText(self.console_window, wrap=tk.WORD)
        self.textbox.pack(expand=True, fill='both')
        self.textbox.configure(state='disabled')  # Делаем TextBox только для чтения

        self.clear_button = tk.Button(self.console_window, text="Очистить", command=self.clear_textbox)
        self.clear_button.pack()

        self.stdout_redirector = TextRedirector(self.textbox, update_status_callback, 'stdout')
        self.stderr_redirector = TextRedirector(self.textbox, update_status_callback, 'stderr')
        sys.stdout = self.stdout_redirector
        sys.stderr = self.stderr_redirector

    def clear_textbox(self):
        self.textbox.configure(state='normal')  # Временно включаем редактирование
        self.textbox.delete(1.0, tk.END)
        self.textbox.configure(state='disabled')  # Отключаем редактирование

    def show_console(self):
        self.console_window.deiconify()
        self.console_window.lift()

    def hide_console(self):
        self.console_window.withdraw()


_console_logger_app_instance = None


def initialize_console_logger(root, update_status_callback=None):
    global _console_logger_app_instance
    if _console_logger_app_instance is None:
        _console_logger_app_instance = ConsoleLoggerApp(root, update_status_callback)
    return _console_logger_app_instance
