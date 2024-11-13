import win32com.client
import tkinter as tk
from tkinter import messagebox, scrolledtext
import json
import os

class OutlookFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Фильтр почты Outlook")
        self.root.geometry("600x500")  # Устанавливаем размер окна

        # Инициализация переменных для фильтров
        self.filters = []
        self.load_filters()

        # Метки и поля для ввода
        self.label_filter = tk.Label(root, text="Введите фильтр (например, отправитель или тема):")
        self.label_filter.grid(row=0, column=0, padx=10, pady=10)

        self.filter_entry = tk.Entry(root, width=50)
        self.filter_entry.grid(row=0, column=1, padx=10, pady=10)

        self.text_search = tk.Label(root, text="Введите текст для поиска (опционально):")
        self.text_search.grid(row=1, column=0, padx=10, pady=10)

        self.text_entry = tk.Entry(root, width=50)
        self.text_entry.grid(row=1, column=1, padx=10, pady=10)

        # Кнопка для фильтрации
        self.filter_button = tk.Button(root, text="Фильтровать", command=self.filter_emails)
        self.filter_button.grid(row=2, column=0, columnspan=2, pady=10)

        # Поле для вывода результатов
        self.result_text = scrolledtext.ScrolledText(root, height=10, width=70, wrap=tk.WORD)
        self.result_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

        # Информация о статусе
        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.grid(row=4, column=0, columnspan=2, pady=5)

        # Если есть фильтры, создаем выпадающий список
        if self.filters:
            self.selected_filter = tk.StringVar(root)
            self.selected_filter.set(self.filters[0])  # Устанавливаем первый фильтр как выбранный по умолчанию
            self.filter_dropdown = tk.OptionMenu(root, self.selected_filter, *self.filters)
            self.filter_dropdown.grid(row=5, column=0, columnspan=2, pady=5)
        else:
            self.filter_dropdown = None  # Если фильтров нет, не создаем выпадающий список

        # Кнопка для сохранения фильтра
        self.save_filter_button = tk.Button(root, text="Сохранить фильтр", command=self.save_filter)
        self.save_filter_button.grid(row=6, column=0, columnspan=2, pady=10)

        # Контейнер для кнопок
        self.canvas = tk.Canvas(root)
        self.scrollbar = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.grid(row=7, column=0, columnspan=2, padx=10, pady=10)
        self.canvas.config(yscrollcommand=self.scrollbar.set)

        self.frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.config(scrollregion=self.canvas.bbox("all")))

        self.scrollbar.grid(row=7, column=2, sticky='ns')  # Прокрутка теперь используется с grid

        self.email_buttons = []  # Список для хранения кнопок для писем

    def load_filters(self):
        # Загружаем фильтры из файла, если он существует
        if os.path.exists("filters.json"):
            with open("filters.json", "r", encoding="utf-8") as f:
                self.filters = json.load(f)

    def save_filters(self):
        # Сохраняем список фильтров в файл
        with open("filters.json", "w", encoding="utf-8") as f:
            json.dump(self.filters, f, ensure_ascii=False)

    def save_filter(self):
        # Сохраняем текущий фильтр
        filter_name = self.filter_entry.get().strip()
        if filter_name and filter_name not in self.filters:
            self.filters.append(filter_name)
            self.save_filters()

            # Обновляем выпадающий список, если фильтры были изменены
            if self.filter_dropdown:
                self.filter_dropdown.destroy()
            self.selected_filter = tk.StringVar(self.root)
            self.selected_filter.set(self.filters[0])  # Устанавливаем первый фильтр как выбранный
            self.filter_dropdown = tk.OptionMenu(self.root, self.selected_filter, *self.filters)
            self.filter_dropdown.grid(row=5, column=0, columnspan=2, pady=5)

    def filter_emails(self):
        filter_text = self.filter_entry.get().strip()
        additional_text = self.text_entry.get().strip()

        if not filter_text:
            messagebox.showwarning("Ошибка", "Пожалуйста, введите фильтр!")
            return

        self.status_label.config(text="Поиск сообщений...")  # Уведомление о начале поиска
        self.root.update()  # Обновляем интерфейс

        # Подключение к Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 - входящие
            messages = inbox.Items

            # Сортировка сообщений по времени
            messages.Sort("[ReceivedTime]", True)

            filtered_emails = []
            self.email_buttons.clear()  # Очищаем старые кнопки

            for message in messages:
                try:
                    if (filter_text.lower() in message.Subject.lower() or filter_text.lower() in message.SenderName.lower()) and \
                            (additional_text.lower() in message.Body.lower() or not additional_text):
                        email_info = f"Тема: {message.Subject}\nОтправитель: {message.SenderName}\nДата: {message.ReceivedTime}"
                        filtered_emails.append(email_info)

                        # Создаем кнопку для открытия письма
                        button = tk.Button(self.frame, text=f"Открыть письмо: {message.Subject}", command=lambda m=message: self.open_email(m))
                        self.email_buttons.append(button)
                except AttributeError:
                    continue  # Игнорировать сообщения, которые не имеют атрибутов

            # Вывод результатов
            self.result_text.delete(1.0, tk.END)
            if filtered_emails:
                self.result_text.insert(tk.END, "\n\n".join(filtered_emails))
            else:
                self.result_text.insert(tk.END, "Нет сообщений, соответствующих фильтру.")

            # Располагаем кнопки
            self.place_buttons()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к Outlook: {e}")
        finally:
            self.status_label.config(text="Поиск завершен.")  # Обновляем статус

    def place_buttons(self):
        # Расположим кнопки с возможностью прокрутки
        for button in self.email_buttons:
            button.grid(row=self.email_buttons.index(button), column=0, pady=5)

    def open_email(self, message):
        # Открываем сообщение в Outlook
        try:
            message.Display()  # Открыть сообщение в Outlook
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть письмо: {e}")

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookFilterApp(root)
    root.mainloop()
