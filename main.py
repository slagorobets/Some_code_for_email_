import imaplib
import email
import os
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from tkcalendar import DateEntry
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.header import decode_header
import base64
import threading
import time
from threading import Semaphore
import babel.numbers
class EmailImporterApp:
    total_emails_processed = 1  # Статическая переменная для сквозного счетчика
    saving_semaphore = Semaphore(1)
    def __init__(self, root):
        # Диапазон дат или полная выборка писем
        def change_check_period(*args):
            if self.date_period.get():
                if self.end_date.state() != (): self.end_date.drop_down()
                if self.start_date.state() != (): self.start_date.drop_down()
                self.label_from.pack_forget()
                self.start_date.pack_forget()
                self.label_to.pack_forget()
                self.end_date.pack_forget()
            else:
                self.label_from.pack(after=self.check_date_period)
                self.start_date.pack(after=self.label_from)
                self.label_to.pack(after=self.start_date)
                self.end_date.pack(after=self.label_to)
        self.root = root
        self.root.title("Импорт электронной почты")
        self.root.geometry("400x600")
        self.save_attachments_var = tk.BooleanVar()
        self.save_attachments_var.set(True)  # По умолчанию сохранять вложения
        self.date_period = tk.BooleanVar()
        self.date_period.set(True)  # По умолчанию выгружать все письма
        style = ttk.Style()
        style.theme_use("clam")
        self.label_email = ttk.Label(root, text="Адрес электронной почты:", font=("Helvetica", 12, "bold"))
        self.label_email.pack(pady=5)
        self.entry_email = ttk.Entry(root, width=40, font=("Helvetica", 10))
        self.entry_email.pack(pady=5)
        self.label_password = ttk.Label(root, text="Пароль:", font=("Helvetica", 12, "bold"))
        self.label_password.pack(pady=5)
        self.entry_password = ttk.Entry(root, width=40, font=("Helvetica", 10), show="*")
        self.entry_password.pack(pady=5)
        self.label_imap = ttk.Label(root, text="IMAP-сервер:", font=("Helvetica", 12, "bold"))
        self.label_imap.pack(pady=5)
        imap_servers = ['imap.mail.ru', 'imap.yandex.ru', 'imap.gmail.com', 'imap.mail.yahoo.com', 'imap-mail.outlook.com']
        self.combo_imap = ttk.Combobox(root, values=imap_servers, width=40, font=("Helvetica", 10))
        self.combo_imap.pack(pady=5)
        self.combo_imap.set('Выберите IMAP сервер или введите здесь')
        self.label_period = ttk.Label(root, text="Период выборки", font=("Helvetica", 12, "bold"))
        self.label_period.pack(pady=5)
        self.check_date_period = ttk.Checkbutton(root, text="Выгрузка за весь период", variable=self.date_period)
        self.date_period.trace_add('write', change_check_period)
        self.check_date_period.pack(pady=5)
        self.label_from = ttk.Label(root, text="c", font=("Helvetica", 8, "bold"))
        self.start_date = DateEntry(root, year=datetime.today().year - 4, day=1, month=1)
        self.label_to = ttk.Label(root, text="по", font=("Helvetica", 8, "bold"))
        self.end_date = DateEntry(root)
        self.check_save_attachments = ttk.Checkbutton(root, text="Сохранять с вложениями", variable=self.save_attachments_var)
        self.check_save_attachments.pack(pady=5)
        self.text_output = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=40, height=10, font=("Helvetica", 10))
        self.text_output.pack(pady=10)
        self.button_import = ttk.Button(root, text="Импортировать", command=self.start_import, style="Accent.TButton")
        self.button_import.pack(pady=10)
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", foreground="white", background="#5E9FFF", font=("Helvetica", 12, "bold"))
        # Создаем директорию для вложений
        self.attachment_dir = os.path.join(os.getcwd(), "email_attachments")
        if not os.path.exists(self.attachment_dir):
            os.makedirs(self.attachment_dir)
    def start_import(self):
        # Сбросить счетчик перед началом новой сессии импорта
        self.total_emails_processed = 1
        # Получение значений из полей ввода
        email_address = self.entry_email.get()
        password = self.entry_password.get()
        imap_server = self.combo_imap.get()
        check_all = self.date_period.get()
        start_date = DateEntry.get_date(self.start_date)
        end_date = DateEntry.get_date(self.end_date)
        # Получение значения флага сохранения вложений
        save_attachments = self.save_attachments_var.get()
        # Проверка заполнения всех полей
        if not email_address or not password or not imap_server:
            messagebox.showwarning("Внимание", "Пожалуйста, введите все необходимые данные.")
            return
        if (start_date - end_date).days >= 0:
            messagebox.showwarning("Внимание", "Пожалуйста, введите правильный временной период.")
            return
        # Создаем уникальное имя папки для текущей сессии импорта
        session_folder_name = f"{email_address}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        session_folder_name = self.replace_invalid_chars(session_folder_name)
        session_folder_path = os.path.join(self.attachment_dir, session_folder_name)
        if not os.path.exists(session_folder_path):
            os.makedirs(session_folder_path)
        # Создаем уникальное имя лог-файла с временной меткой
        log_filename = f"email_log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
        log_filepath = os.path.join(os.getcwd(), log_filename)
        # Получаем общее количество писем в папке
        total_emails_in_folder = sum(len(self.get_email_uids(email_address, password, imap_server, folder, check_all, start_date, end_date)) for folder in self.get_folders(email_address, password, imap_server))
        # Создаем окно прогресса и запускаем импорт в отдельном потоке
        progress_window = ProgressWindow(self.root, total_emails_in_folder)
        progress_thread = threading.Thread(target=self.import_emails_async, args=(email_address, password, imap_server, check_all, start_date, end_date, save_attachments, session_folder_path, log_filepath, progress_window))
        progress_thread.start()
    def import_emails_async(self, email_address, password, imap_server, check_all, start_date, end_date, save_attachments, session_folder_path, log_filepath, progress_window):
        progress_window.show()
        folders = self.get_folders(email_address, password, imap_server)
        total_emails = sum(len(self.get_email_uids(email_address, password, imap_server, folder, check_all, start_date, end_date)) for folder in folders)
        # Создаем список для хранения всех запущенных потоков
        thread_list = []
        with ThreadPoolExecutor(max_workers=9) as executor:
            try:
                # Захватываем семафор перед началом сохранения писем
                self.saving_semaphore.acquire()
                for folder in folders:
                    folder_path = os.path.join(session_folder_path, self.decode_folder_name(folder))
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    uids = self.get_email_uids(email_address, password, imap_server, folder, check_all, start_date, end_date)
                    for uid in uids:
                        thread = executor.submit(self.import_emails, email_address, password, imap_server, folder, save_attachments, folder_path, log_filepath, uid, progress_window, total_emails)
                        thread_list.append(thread)

                        # Небольшая задержка для имитации работы
                        time.sleep(0.1)
            except Exception as e:
                print(f"Ошибкапри импорте писем: {e}")
                # В случае ошибки освобождаем семафор
                self.saving_semaphore.release()
                return
            finally:
                # Освобождаем семафор после завершения всех потоков
                self.saving_semaphore.release()
        # Дождемся завершения всех потоков
        for thread in thread_list:
            thread.result()
        # Закрываем окно прогресса после завершения
        progress_window.close()
        # Выводим сообщение о завершении импорта
        self.after_completion_message()
    def import_emails(self, email_address, password, imap_server, folder, save_attachments, folder_path, log_filepath, uid, progress_window, total_emails):
        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            mail.select(mail._quote(folder))

            result, msg_data = mail.uid('fetch', uid, '(RFC822)')
            if result == 'OK':
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                # Получаем информацию о письме
                from_, to, subject, email_date, email_time, num_attachments, msg_size = self.get_email_info(msg, save_attachments, folder_path, uid)
                # Получаем путь к файлу .eml
                msg_filepath = os.path.join(folder_path, f"{uid}_{email_date}_{email_time}.eml")
                # Записываем информацию в лог
                self.write_to_log(log_filepath, uid, email_date, email_time, from_, to, subject, num_attachments, msg_size, msg_filepath, folder)
                self.total_emails_processed += 1
                # Обновляем прогресс в окне
                progress_window.update_progress(self.total_emails_processed, total_emails)
        except Exception as e:
            print(f"Ошибка обработки письма {uid}: {e}")
        finally:
            mail.logout()
    def get_folders(self, email_address, password, imap_server):
        folders = []
        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            folders_response, folders_data = mail.list()
            if folders_response == 'OK':
                if folders_data:
                    for folder_info in folders_data:
                        # Разбиваем строку, чтобы получить закодированное имя папки
                        split_string = ' "' + folder_info.decode().split('"')[1] + '" '
                        folder_name = folder_info.decode().split(split_string)[1].strip('"')
                        folders.append(folder_name)
            mail.logout()
        except Exception as e:
            print(f"Ошибка получения списка папок: {e}")
        return folders

    def get_email_uids(self, email_address, password, imap_server, folder, check_all, start_date, end_date):
        uids = []
        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            mail.select(mail._quote(folder))
            if check_all:
                result, search_data = mail.uid('search', None, 'ALL')
            else:
                start_date = start_date.strftime("%d-%b-%Y")
                end_date = end_date.strftime("%d-%b-%Y")
                result, search_data = mail.uid('search', None, 'SINCE {start_date} BEFORE {end_date}'.format(start_date=start_date, end_date=end_date))
            if result == 'OK':
                uids = search_data[0].split()
                # Сортировка UID в обратном порядке
                uids = sorted(uids, key=int, reverse=True)
        except Exception as e:
            print(f"Ошибка при получении UID: {e}")
        finally:
            mail.logout()
        return uids
    def save_email_to_eml(self, msg, folder_path, email_number, email_date, email_time):
        encodings = ['utf-8','koi8-r', 'latin1', 'utf-16', 'iso-8859-1', 'windows-1251', 'utf-8-sig', 'replace', 'ignore', 'backslashreplace', 'xmlcharrefreplace', 'unicode_escape']
        for encoding in encodings:
            try:
                # Создаем пустой EML-файл
                eml_filepath = os.path.join(folder_path, f"{int(email_number)}_{email_date}_{email_time}.eml")
                with open(eml_filepath, "w", encoding=encoding) as f:
                    # Преобразуем объект сообщения в строку и записываем в EML-файл
                    f.write(msg.as_string())
                print(f"Сообщение {int(email_number)} сохранено в формате .eml с кодировкой {encoding}")
                # Возвращаем размер сохраненного сообщения
                return os.path.getsize(eml_filepath)
            except Exception as e:
                print(f"Ошибка при сохранении сообщения {int(email_number)} с кодировкой {encoding}: {e}")
        # Если все попытки сохранения с указанием кодировки завершились неудачно, попробуем сохранить без указания кодировки
        try:
            # Создаем пустой EML-файл без указания кодировки
            eml_filepath = os.path.join(folder_path, f"{int(email_number)}_{email_date}_{email_time}.eml")
            with open(eml_filepath, "w") as f:
                # Преобразуем объект сообщения в строку и записываем в EML-файл
                f.write(msg.as_string())
            print(f"Сообщение {int(email_number)} сохранено в формате .eml без указания кодировки")
            # Возвращаем размер сохраненного сообщения
            return os.path.getsize(eml_filepath)
        except Exception as e:
            print(f"Ошибка при сохранении сообщения {int(email_number)} без указания кодировки: {e}")
        return 0
    def save_email_to_eml_without_attachments(self, msg, folder_path, email_number, email_date, email_time):
        pass
    def write_to_log(self, log_filepath, email_number, email_date, email_time, from_, to, subject, num_attachments, msg_size, msg_filepath, folder_name):
        try:
            formatted_date = datetime.strptime(email_date, "%d-%m-%Y").strftime("%d.%m.%Y")
            formatted_time = email_time.replace('-', ':')
            folder_name = self.decode_folder_name(folder_name)
            email_number_int = int(email_number)  # Преобразование номера письма в число
            with open(log_filepath, "a") as f:
                f.write(f'"{self.total_emails_processed}";"{folder_name}";"{email_number_int}";"{formatted_date}";"{formatted_time}";"{from_}";"{to}";"{subject}";"{num_attachments}";"{msg_size}";"{msg_filepath}";\n')
        except Exception as e:
            print(f"Ошибка при записи в журнал: {e}")
    def get_email_info(self, msg, save_attachments, folder_path, uid):
        from_, to = self.get_email_sender_and_receiver(msg)
        subject = self.clean_subject(msg.get("Subject", ""))
        email_date = email.utils.parsedate_to_datetime(msg.get("Date", ""))
        if email_date:
            email_time = email_date.strftime("%H-%M-%S")
            email_date = email_date.strftime("%d-%m-%Y")
        else:
            email_date = email_time = ""
        # Проверяем наличие вложений только если флаг save_attachments установлен в True
        num_attachments = self.count_attachments(msg) if save_attachments else 0
        # Вызываем соответствующую функцию сохранения для определения размера сообщения
        if save_attachments:
            msg_size = self.save_email_to_eml(msg, folder_path, uid, email_date, email_time)
        else:
            msg_size = self.save_email_to_eml_without_attachments(msg, folder_path, uid, email_date, email_time)

        return from_, to, subject, email_date, email_time, num_attachments, msg_size
    def clean_subject(self, subject):
        """Очистка темы письма от инородных символов."""
        decoded_subject, encoding = decode_header(subject)[0]
        if isinstance(decoded_subject, bytes):
            decoded_subject = decoded_subject.decode(encoding or 'utf-8', errors='replace')
        return re.sub(r'[^\w\s]', '', decoded_subject)
    def get_email_sender_and_receiver(self, msg):
        sender_email = None
        receiver_email = None
        sender = msg.get("From")
        sender_email = self.extract_email_from_header(sender)
        receiver = msg.get("To")
        receiver_email = self.extract_email_from_header(receiver)
        if sender_email is None:
            sender_email = sender
        return sender_email, receiver_email
    def extract_email_from_header(self, header):
        if header:
            decoded_header, encoding = decode_header(header)[0]
            if isinstance(decoded_header, bytes):
                decoded_header = decoded_header.decode(encoding or 'utf-8', errors='replace')
            match_header = re.search(r'[\w\.-]+@[\w\.-]+', header)
            match = re.search(r'[\w\.-]+@[\w\.-]+', decoded_header)
            if match_header:
                return match_header.group(0)
            else:
                if match:
                    return match.group(0)
                else:
                    return decoded_header
        return None
    def count_attachments(self, msg):
        num_attachments = 0
        for part in msg.walk():
            if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
                num_attachments += 1
        return num_attachments
    def after_completion_message(self):
        messagebox.showinfo("Завершено", "Импорт почты завершен успешно.")
    def replace_invalid_chars(self, name):
        """Заменяет недопустимые символы в имени файла или папки."""
        return re.sub(r'[<>:"/\\|?*]', '_', name)
    def decode_folder_name(self, folder):
        """Декодирование имени папки, закодированного в соответствии с IMAP UTF7."""
        lst = folder.split('&')
        out = lst[0]
        for e in lst[1:]:
            u, a = e.split('-', 1)  # u: utf16 между & и 1-м -, a: ASCII символы после него
            if u == '':
                out += '&'
            else:
                out += self.b64padanddecode(u)
            out += self.replace_invalid_chars(a)  # Обработка недопустимых символов в ASCII части
        out = self.replace_invalid_chars(out)  # Обработка недопустимых символов в итоговом имени папки
        return out
    @staticmethod
    def b64padanddecode(b):
        """Декодирование незаполненных данных base64."""
        b += (-len(b) % 4) * '='  # дополнение base64 (если добавить '===' , все равно не будет корректного дополнения)
        return base64.b64decode(b.encode('ascii'), altchars='+,', validate=True).decode('utf-16-be')
    def after_completion_message(self):
        """Выводит сообщение об успешном завершении импорта, включая количество скачанных писем."""
        messagebox.showinfo("Импорт завершен", f"Импорт писем успешно завершен. Скачано писем: {self.total_emails_processed - 1}")
class ProgressWindow:
    def __init__(self, root, total):
        self.root = root
        self.progress_window = tk.Toplevel(root)
        self.progress_window.title("Прогресс выполнения")
        self.progress_label = ttk.Label(self.progress_window, text="Выполняется импорт...")
        self.progress_label.pack(pady=10)
        self.progress_bar = ttk.Progressbar(self.progress_window, length=300, mode="determinate")
        self.progress_bar.pack(pady=10)
        self.progress_percent_label = ttk.Label(self.progress_window, text="")
        self.progress_percent_label.pack()
        self.total = total  # Добавляем общее количество сообщений
    def show(self):
        self.root.update()
        self.progress_window.grab_set()
        self.progress_window.focus_set()
    def update_progress(self, current, total):
        progress_percentage = 0
        progress_percentage = (current / total) * 100  # Вычисляем процент выполнения
        self.progress_bar["value"] = progress_percentage
        self.progress_percent_label["text"] = f"Прогресс: {int(progress_percentage)}%"
        self.root.update()
    def close(self):
        self.progress_window.destroy()
def main():
    root = tk.Tk()
    EmailImporterApp(root)
    root.mainloop()
if __name__ == "__main__":
    main()