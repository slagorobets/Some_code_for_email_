import imaplib
import email
import os
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.header import decode_header
import base64


class EmailImporterApp:
    total_emails_processed = 1  # Статическая переменная для сквозного счетчика

    def __init__(self, root):
        self.root = root
        self.root.title("Импорт электронной почты")
        self.root.geometry("400x600")

        self.save_attachments_var = tk.BooleanVar()
        self.save_attachments_var.set(True)  # По умолчанию сохранять вложения

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

        self.entry_imap = ttk.Entry(root, width=40, font=("Helvetica", 10))
        self.entry_imap.pack(pady=5)

        self.label_period = ttk.Label(root, text="Период (за последние ? лет):", font=("Helvetica", 12, "bold"))
        self.label_period.pack(pady=5)

        self.entry_period = ttk.Entry(root, width=10, font=("Helvetica", 10))
        self.entry_period.pack(pady=5)

        self.check_save_attachments = ttk.Checkbutton(root, text="Сохранять с вложениями",
                                                      variable=self.save_attachments_var)
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
        # Получение значений из полей ввода
        email_address = self.entry_email.get()
        password = self.entry_password.get()
        imap_server = self.entry_imap.get()
        period = self.entry_period.get()

        # Проверка периода
        if period.strip() and period.isdigit():
            period = int(period)
        else:
            period = None

        # Получение значения флага сохранения вложений
        save_attachments = self.save_attachments_var.get()

        # Проверка заполнения всех полей
        if not email_address or not password or not imap_server:
            messagebox.showwarning("Внимание", "Пожалуйста, введите все необходимые данные.")
            return

        # Создаем уникальное имя папки для текущей сессии импорта
        session_folder_name = f"{email_address}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        session_folder_name = self.replace_invalid_chars(session_folder_name)
        session_folder_path = os.path.join(self.attachment_dir, session_folder_name)
        if not os.path.exists(session_folder_path):
            os.makedirs(session_folder_path)

        # Создаем уникальноеимя лог-файла с временной меткой
        log_filename = f"email_log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
        log_filepath = os.path.join(os.getcwd(), log_filename)

        # Создаем список для хранения всех запущенных потоков
        thread_list = []

        with ThreadPoolExecutor(max_workers=5) as executor:
            try:
                folders = self.get_folders(email_address, password, imap_server)
                for folder in folders:
                    folder_path = os.path.join(session_folder_path, self.decode_folder_name(folder))
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    uids = self.get_email_uids(email_address, password, imap_server, folder, period)
                    for uid in uids:
                        thread = executor.submit(self.import_emails, email_address, password, imap_server, folder,
                                                 period, save_attachments, folder_path, log_filepath, uid)
                        thread_list.append(thread)
            except Exception as e:
                print(f"Ошибка при импорте писем: {e}")

        # Дождемся завершения всех потоков
        for thread in thread_list:
            thread.result()

        # Выводим сообщение о завершении импорта
        self.after_completion_message()

    def import_emails(self, email_address, password, imap_server, folder, period, save_attachments, folder_path,
                      log_filepath, uid):
        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            mail.select(folder)

            result, msg_data = mail.uid('fetch', uid, '(RFC822)')
            if result == 'OK':
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                # Получаем информацию о письме
                from_, to, subject, email_date, email_time, num_attachments, msg_size = self.get_email_info(msg,
                                                                                                            save_attachments,
                                                                                                            folder_path,
                                                                                                            uid)
                # Получаем путь к файлу .msg
                msg_filepath = os.path.join(folder_path, f"{uid}_{email_date}_{email_time}.msg")
                # Записываем информацию в лог
                self.write_to_log(log_filepath, uid, email_date, email_time, from_, to, subject, num_attachments,
                                  msg_size, msg_filepath)
                self.total_emails_processed += 1
        except Exception as e:
            print(f"Ошибка обработки письма {uid}: {e}")

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
                        folders.append(folder_name.replace('/', "\"))
            mail.logout()
        except Exception as e:
            print(f"Ошибка получения списка папок: {e}")
        return folders

    def get_email_uids(self, email_address, password, imap_server, folder, period):
        uids = []
        try:
            mail = imaplib.IMAP4_SSL(imap_server)
            mail.login(email_address, password)
            mail.select(folder)
            if period:
                date = (datetime.now() - timedelta(days=365 * period)).strftime("%d-%b-%Y")
                result, search_data = mail.uid('search', None, f'(SENTSINCE {date})')
            else:
                result, search_data = mail.uid('search', None, 'ALL')
            if result == 'OK':
                uids = search_data[0].split()
        except Exception as e:
            print(f"Ошибка при получении UID: {e}")
        finally:
            mail.logout()
        return uids

    def save_email_to_msg(self, msg, folder_path, email_number, email_date, email_time):
        try:
            # Сохраняем сообщение
            msg_bytes = msg.as_bytes()
            msg_filepath = os.path.join(folder_path, f"{email_number}_{email_date}_{email_time}.msg")
            with open(msg_filepath, "wb") as f:
                f.write(msg_bytes)

            print(f"Сообщение {email_number} сохранено в формате .msg")

            # Возвращаем размер сохраненного сообщения
            return len(msg_bytes)
        except Exception as e:
            print(f"Ошибка при сохранении сообщения {email_number}: {e}")
            return 0

    def save_email_to_msg_without_attachments(self, msg, folder_path, email_number, email_date, email_time):
        try:
            # Удаляем вложения из сообщения
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                filename = part.get_filename()
                if filename:
                    part.set_payload(None)

            # Сохраняем сообщение без вложений
            msg_bytes = msg.as_bytes()
            msg_filepath = os.path.join(folder_path, f"{email_number}_{email_date}_{email_time}.msg")
            with open(msg_filepath, "wb") as f:
                f.write(msg_bytes)

            print(f"Сообщение {email_number} сохранено в формате .msg без вложений")

            # Возвращаем размер сохраненного сообщения
            return len(msg_bytes)
        except Exception as e:
            print(f"Ошибка при сохранении сообщения {email_number}: {e}")
            return 0

    def write_to_log(self, log_filepath, email_number, email_date, email_time, from_, to, subject, num_attachments,
                     msg_size, msg_filepath):
        try:
            formatted_date = datetime.strptime(email_date, "%d-%m-%Y").strftime("%d.%m.%Y")
            formatted_time = email_time.replace('-', ':')
            email_number_int = int(email_number)  # Преобразование номера письма в число
            with open(log_filepath, "a") as f:
                f.write(
                    f"{self.total_emails_processed};{email_number_int};{formatted_date};{formatted_time};{from_};{to};{subject};{num_attachments};{msg_size};{msg_filepath};\n")
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
            msg_size = self.save_email_to_msg(msg, folder_path, uid, email_date, email_time)
        else:
            msg_size = self.save_email_to_msg_without_attachments(msg, folder_path, uid, email_date, email_time)

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
        """Декодирование имени папки, закодированного в соответствии с IMAP

 UTF7."""
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


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailImporterApp(root)
    root.mainloop()