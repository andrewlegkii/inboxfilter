import win32com.client
import tkinter as tk
from tkinter import messagebox, scrolledtext

class OutlookFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Фильтр почты Outlook")
        self.root.geometry("600x400")  # Устанавливаем размер окна

        # Метки и поля для ввода
        self.label_filter = tk.Label(root, text="Введите фильтр (например, отправитель или тема):")
        self.label_filter.grid(row=0, column=0, padx=10, pady=10)

        self.filter_entry = tk.Entry(root, width=50)
        self.filter_entry.grid(row=0, column=1, padx=10, pady=10)

        # Кнопка для фильтрации
        self.filter_button = tk.Button(root, text="Фильтровать", command=self.filter_emails)
        self.filter_button.grid(row=1, column=0, columnspan=2, pady=10)

        # Поле для вывода результатов
        self.result_text = scrolledtext.ScrolledText(root, height=15, width=70, wrap=tk.WORD)
        self.result_text.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

        # Информация о статусе
        self.status_label = tk.Label(root, text="", fg="blue")
        self.status_label.grid(row=3, column=0, columnspan=2, pady=5)

    def filter_emails(self):
        filter_text = self.filter_entry.get().strip()
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
            for message in messages:
                try:
                    if filter_text.lower() in message.Subject.lower() or filter_text.lower() in message.SenderName.lower():
                        filtered_emails.append(f"Тема: {message.Subject}\nОтправитель: {message.SenderName}\nДата: {message.ReceivedTime}\n")
                except AttributeError:
                    continue  # Игнорировать сообщения, которые не имеют атрибутов

            # Вывод результатов
            self.result_text.delete(1.0, tk.END)
            if filtered_emails:
                self.result_text.insert(tk.END, "\n\n".join(filtered_emails))
            else:
                self.result_text.insert(tk.END, "Нет сообщений, соответствующих фильтру.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось подключиться к Outlook: {e}")
        finally:
            self.status_label.config(text="Поиск завершен.")  # Обновляем статус

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = OutlookFilterApp(root)
    root.mainloop()
