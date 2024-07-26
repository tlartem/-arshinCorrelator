import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
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
    def __init__(self, parent, update_status_callback=None):
        self.console_frame = ttk.Frame(parent)
        self.console_frame.pack(fill=tk.BOTH, expand=True)

        self.textbox = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, state='disabled')
        self.textbox.pack(expand=True, fill='both', padx=10, pady=10)

        self.stdout_redirector = TextRedirector(self.textbox, update_status_callback, 'stdout')
        self.stderr_redirector = TextRedirector(self.textbox, update_status_callback, 'stderr')
        sys.stdout = self.stdout_redirector
        sys.stderr = self.stderr_redirector

        # Добавляем кнопки управления
        self.button_frame = ttk.Frame(self.console_frame)
        self.button_frame.pack(fill=tk.X, pady=5, padx=10)

        self.clear_button = ttk.Button(self.button_frame, text="Очистить", command=self.clear_console)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.copy_button = ttk.Button(self.button_frame, text="Копировать", command=self.copy_to_clipboard)
        self.copy_button.pack(side=tk.LEFT, padx=5)

    def show_console(self):
        self.console_frame.pack(fill=tk.BOTH, expand=True)
        self.console_frame.lift()

    def hide_console(self):
        self.console_frame.pack_forget()

    def clear_console(self):
        self.textbox.configure(state='normal')
        self.textbox.delete('1.0', tk.END)
        self.textbox.configure(state='disabled')

    def copy_to_clipboard(self):
        self.console_frame.clipboard_clear()
        self.console_frame.clipboard_append(self.textbox.get('1.0', tk.END))
        messagebox.showinfo("Копировать", "Содержимое консоли скопировано в буфер обмена")

    def write_message(self, message):
        self.stdout_redirector.write(message)


_console_logger_app_instance = None


def initialize_console_logger(parent, update_status_callback=None):
    global _console_logger_app_instance
    if _console_logger_app_instance is None:
        _console_logger_app_instance = ConsoleLoggerApp(parent, update_status_callback)
    return _console_logger_app_instance
