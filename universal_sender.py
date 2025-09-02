#!/usr/bin/env python3
# -*- coding: utf-8 -*- 

import asyncio
import smtplib
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
import pandas as pd
import logging
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pyrogram import Client
from pyrogram.errors import (
    SessionPasswordNeeded,
    PhoneNumberInvalid,
    PhoneCodeInvalid, 
    PhoneCodeExpired,
    PasswordHashInvalid,
    FloodWait,
    BadRequest
)
import threading
import os
import json
import math
import random
import colorsys
import mimetypes
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('universal_sender.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class ThemeManager:
    """Управление темами интерфейса"""
    
    def __init__(self, root):
        self.root = root
        self.is_dark = False
        
        # Цвета для светлой темы
        self.light_theme = {
            'bg': '#ffffff',
            'fg': "#0A0A0A",
            'select_bg': '#0078d4',
            'select_fg': '#ffffff',
            'entry_bg': '#ffffff',
            'entry_fg': "#000000",
            'button_bg': '#f0f0f0',
            'frame_bg': '#f5f5f5',
            'text_bg': '#ffffff',
            'disabled_fg': '#666666',
            'log_bg': '#f8f8f8'
        }
        
        # Цвета для темной темы
        self.dark_theme = {
            'bg': "#4d4c4c",
            'fg': '#e0e0e0', 
            'select_bg': '#404040',
            'select_fg': '#ffffff',
            'entry_bg': '#333333',
            'entry_fg': '#ffffff',
            'button_bg': '#404040',
            'frame_bg': '#1e1e1e',     # Убираем светлые фреймы - тот же цвет что и bg
            'text_bg': '#2a2a2a',
            'disabled_fg': '#cccccc',
            'log_bg': '#252525'
        }
    
    def toggle_theme(self):
        """Переключение темы"""
        self.is_dark = not self.is_dark
        self.apply_theme()
        return "☀️ Светлая" if self.is_dark else "🌙 Темная"
    
    def get_current_theme(self):
        """Получить текущую тему"""
        return self.dark_theme if self.is_dark else self.light_theme
    
    def apply_theme(self):
        """Применить тему ко всем элементам"""
        theme = self.get_current_theme()
        
        # Применяем тему к корневому окну
        self.root.configure(bg=theme['bg'])
        
        # Настройка стиля для ttk виджетов
        style = ttk.Style()
        
        if self.is_dark:
            style.theme_use('clam')
            
            # Основные элементы
            style.configure('TLabel', 
                        background=theme['bg'], 
                        foreground=theme['fg'])
            
            style.configure('TFrame', 
                        background=theme['bg'])
            
            # КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ: настройка LabelFrame
            style.configure('TLabelFrame', 
                        background=theme['bg'],        # Темный фон
                        foreground=theme['fg'],        # Светлый текст заголовка
                        relief='flat',
                        borderwidth=0,                 # Убираем границы
                        labeloutside=False)            # Заголовок внутри
            
            # Настройка внутренней области LabelFrame
            style.configure('TLabelFrame.Label',
                        background=theme['bg'],        # Темный фон под заголовком
                        foreground=theme['fg'])        # Светлый текст заголовка
            
            style.configure('TButton', 
                        background=theme['button_bg'], 
                        foreground=theme['fg'],
                        relief='flat',
                        borderwidth=0,
                        focuscolor='none')
            
            style.map('TButton',
                    background=[('active', '#505050'),
                                ('pressed', '#353535')])
            
            style.configure('TEntry', 
                        fieldbackground=theme['entry_bg'], 
                        foreground=theme['entry_fg'],
                        bordercolor=theme['entry_bg'],
                        insertcolor=theme['fg'],
                        relief='flat',
                        borderwidth=1)
            
            style.configure('TCheckbutton', 
                        background=theme['bg'], 
                        foreground=theme['fg'],
                        focuscolor='none')
            
            style.configure('TRadiobutton', 
                        background=theme['bg'], 
                        foreground=theme['fg'],
                        focuscolor='none')
            
            style.configure('TNotebook', 
                        background=theme['bg'],
                        borderwidth=0)
            
            style.configure('TNotebook.Tab', 
                        background=theme['button_bg'], 
                        foreground=theme['fg'],
                        padding=[12, 6],
                        relief='flat')
            
            style.map('TNotebook.Tab',
                    background=[('selected', '#505050')])
                    
            style.configure('TScale',
                        background=theme['bg'],
                        troughcolor=theme['entry_bg'],
                        sliderrelief='flat')
                        
            style.configure('TProgressbar',
                        background='#404040',
                        troughcolor=theme['entry_bg'],
                        borderwidth=0,
                        relief='flat')
        else:
            # Возвращаем стандартную светлую тему
            style.theme_use('default')

class UniversalSenderConfig:
    """Конфигурация приложения"""
    def __init__(self):
        self.config_file = 'config.json'
        self.load_config()
    
    def load_config(self):
        """Загрузка конфигурации из файла"""
        default_config = {
            'telegram': {
                'api_id': 00000000,
                'api_hash': 'PAST_HASH_HERE',
                'session_name': 'universal_sender'
            },
            'email': {
                'provider': 'custom',
                'smtp_server': 'smtp.zoho.eu',
                'smtp_port': 587,
                'use_tls': True,
                'username': 'test-rassilka@zohomail.eu',
                'password': '9GETXTPXaBD9'
            },
            'sending': {
                'delay_seconds': 1.0,
                'send_telegram': True,
                'send_email': True
            },
            'message': {
                'subject': 'Важное сообщение',
                'template': '''Здравствуйте, {name}!

Это тестовое сообщение универсальной рассылки.

Вы получили это сообщение в рамках тестирования системы массовой рассылки.

С уважением,
Команда разработки'''
            },
            'smtp_presets': {
                'gmail': {
                    'smtp_server': 'smtp.gmail.com',
                    'smtp_port': 587,
                    'use_tls': True
                },
                'yandex': {
                    'smtp_server': 'smtp.yandex.ru',
                    'smtp_port': 587,
                    'use_tls': True
                },
                'mailru': {
                    'smtp_server': 'smtp.mail.ru',
                    'smtp_port': 587,
                    'use_tls': True
                },
                'zoho': {
                    'smtp_server': 'smtp.zoho.eu',
                    'smtp_port': 587,
                    'use_tls': True
                }
            }
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Обновляем только существующие ключи
                    for key in default_config:
                        if key in loaded_config:
                            if isinstance(default_config[key], dict):
                                default_config[key].update(loaded_config[key])
                            else:
                                default_config[key] = loaded_config[key]
            except Exception as e:
                logging.warning(f"Ошибка загрузки конфигурации: {e}")
        
        self.config = default_config
        self.save_config()
    
    def save_config(self):
        """Сохранение конфигурации в файл"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Ошибка сохранения конфигурации: {e}")

class DataProcessor:
    """Обработка данных из Excel файлов"""
    
    @staticmethod
    def load_excel_file(file_path):
        """Загрузка Excel файла с правильной обработкой заголовков"""
        try:
            # Сначала читаем без заголовков для анализа структуры
            df_raw = pd.read_excel(file_path, header=None)
            logging.info(f"Прочитано {len(df_raw)} строк из файла")
            
            # Ищем строку с реальными заголовками
            header_row = None
            for i in range(min(5, len(df_raw))):  # Проверяем первые 5 строк
                row_values = df_raw.iloc[i].astype(str).str.lower()
                
                # Ищем характерные заголовки
                keywords = ['username', 'tg', 'telegram', 'name', 'email', 'fio', 'ФИО'.lower()]
                matches = sum(1 for val in row_values if any(keyword in val for keyword in keywords))
                
                if matches >= 2:  # Если найдено 2+ ключевых слова
                    header_row = i
                    logging.info(f"Найдена строка заголовков: {i + 1}")
                    break
            
            if header_row is None:
                # Если не нашли заголовки автоматически, берем первую строку
                logging.warning("Автоматический поиск заголовков не удался, используем первую строку")
                header_row = 0
            
            # Читаем файл с правильными заголовками
            df = pd.read_excel(file_path, header=header_row)
            
            # Очищаем названия столбцов от лишних пробелов
            df.columns = df.columns.astype(str).str.strip()
            
            # Удаляем полностью пустые строки
            df = df.dropna(how='all')
            
            logging.info(f"Загружено {len(df)} записей из файла")
            logging.info("Найденные колонки:")
            for i, col in enumerate(df.columns):
                logging.info(f"  {i+1}. '{col}'")
            
            return df
            
        except Exception as e:
            logging.error(f"Ошибка загрузки файла: {e}")
            raise
    
    @staticmethod
    def process_contacts(df):
        """Обработка контактов для рассылки"""
        try:
            # Ищем колонки по ключевым словам
            telegram_col = DataProcessor.find_column(df, [
                'username', 'tg name', 'tg_name', 'telegram', 'nick', 'Username'
            ])
            
            email_col = DataProcessor.find_column(df, [
                'email', 'mail', 'почта', 'e-mail', 'Email'
            ])
            
            name_col = DataProcessor.find_column(df, [
                'fio', 'имя', 'name', 'ФИО', 'Name', 'FIO', 'TG Name'
            ])
            
            logging.info(f"Найденные колонки: TG='{telegram_col}', Email='{email_col}', Name='{name_col}'")
            
            contacts = []
            
            for index, row in df.iterrows():
                try:
                    contact = {
                        'name': DataProcessor.get_safe_value(row, name_col, f'Контакт_{index+1}'),
                        'telegram': DataProcessor.get_safe_value(row, telegram_col, ''),
                        'email': DataProcessor.get_safe_value(row, email_col, ''),
                        'row_number': index + 1
                    }
                    
                    # Очистка telegram username
                    if contact['telegram']:
                        tg_value = str(contact['telegram']).strip()
                        if tg_value and tg_value.lower() not in ['nan', 'none', '']:
                            # Если это TG ID (только цифры), пропускаем
                            if tg_value.isdigit():
                                contact['telegram'] = ''
                                logging.debug(f"Пропускаем TG ID {tg_value} для {contact['name']}")
                            # Если это username без @, добавляем @
                            elif not tg_value.startswith('@'):
                                contact['telegram'] = '@' + tg_value
                            else:
                                contact['telegram'] = tg_value
                        else:
                            contact['telegram'] = ''
                    
                    # Добавляем контакт если есть хотя бы один способ связи
                    if contact['telegram'] or contact['email']:
                        contacts.append(contact)
                        
                except Exception as row_error:
                    logging.warning(f"Ошибка обработки строки {index + 1}: {row_error}")
                    continue
            
            logging.info(f"Обработано {len(contacts)} контактов для рассылки")
            return contacts
            
        except Exception as e:
            logging.error(f"Ошибка обработки контактов: {e}")
            raise
    
    @staticmethod
    def find_column(df, keywords):
        """Поиск колонки по ключевым словам"""
        # Сначала ищем точное совпадение
        for keyword in keywords:
            for col in df.columns:
                if col.lower() == keyword.lower():
                    return col
        
        # Затем ищем частичное совпадение
        for keyword in keywords:
            for col in df.columns:
                if keyword.lower() in col.lower():
                    return col
        return None
    
    @staticmethod
    def get_safe_value(row, column, default=''):
        """Безопасное получение значения из строки"""
        if column is None:
            return default
        
        value = row.get(column, default)
        if pd.isna(value):
            return default
        
        return str(value).strip()

class TelegramSender:
    """Отправка сообщений через Telegram"""
    
    def __init__(self, config):
        self.config = config['telegram']
        self.app = None
        self.session_file = f"{self.config['session_name']}.session"
    
    def is_session_exists(self):
        """Проверка существования файла сессии"""
        exists = os.path.exists(self.session_file)
        return exists
    
    async def initialize(self):
        """Инициализация Telegram клиента с проверкой сессии"""
        try:
            # Создаем клиент
            self.app = Client(
                self.config['session_name'],
                api_id=self.config['api_id'],
                api_hash=self.config['api_hash']
            )
            
            logging.info("🔄 Подключение к Telegram...")
            await self.app.start()
            
            # Проверяем авторизацию
            me = await self.app.get_me()
            if me:
                logging.info(f"✅ Telegram авторизован как: {me.first_name} (@{me.username or 'без username'})")
                return True
            else:
                logging.error("❌ Не удалось получить информацию о пользователе")
                return False
                
        except Exception as e:
            logging.error(f"❌ Ошибка инициализации Telegram: {e}")
            # Очищаем поврежденную сессию
            if os.path.exists(self.session_file):
                try:
                    os.remove(self.session_file)
                    logging.info("🗑️ Поврежденный файл сессии удален")
                except:
                    pass
            return False
    
    async def send_message(self, username, message):
        """Отправка сообщения в Telegram"""
        try:
            if not self.app:
                logging.error("❌ Telegram клиент не инициализирован")
                return False
                
            if not username or not username.startswith('@'):
                logging.warning(f"❌ Некорректный username: {username}")
                return False

            await self.app.send_message(username, message)
            logging.info(f"✅ Telegram сообщение отправлено: {username}")
            return True

        except Exception as e:
            err = str(e)
            if "USERNAME_NOT_OCCUPIED" in err:
                logging.error(f"❌ Username не существует: {username}")
            elif "PRIVACY" in err or "privacy" in err.lower():
                logging.error(f"⚠️ {username} запретил получать сообщения")
            elif "FLOOD_WAIT" in err:
                logging.error(f"⏳ Flood wait для {username}")
            else:
                logging.error(f"❌ Ошибка отправки в Telegram {username}: {e}")
            return False
        
    async def send_message_with_attachment(self, username, message, attachment_path=None):
        """Отправка сообщения с возможным вложением в Telegram"""
        try:
            if not self.app:
                logging.error("❌ Telegram клиент не инициализирован")
                return False
                
            if not username or not username.startswith('@'):
                logging.warning(f"❌ Некорректный username: {username}")
                return False

            if attachment_path and os.path.exists(attachment_path):
                # Отправляем с файлом
                await self.app.send_document(username, attachment_path, caption=message)
                logging.info(f"✅ Telegram сообщение с файлом отправлено: {username}")
            else:
                # Отправляем только текст
                await self.app.send_message(username, message)
                logging.info(f"✅ Telegram сообщение отправлено: {username}")
            
            return True

        except Exception as e:
            err = str(e)
            if "USERNAME_NOT_OCCUPIED" in err:
                logging.error(f"❌ Username не существует: {username}")
            elif "PRIVACY" in err or "privacy" in err.lower():
                logging.error(f"⚠️ {username} запретил получать сообщения")
            elif "FLOOD_WAIT" in err:
                logging.error(f"⏳ Flood wait для {username}")
            else:
                logging.error(f"❌ Ошибка отправки в Telegram {username}: {e}")
            return False
    
    async def close(self):
        """Закрытие соединения"""
        if self.app:
            try:
                await self.app.stop()
                logging.info("🔌 Telegram соединение закрыто")
            except Exception as e:
                logging.warning(f"⚠️ Ошибка при закрытии Telegram: {e}")

    def get_session_status(self):
        """Получить статус сессии"""
        if not self.is_session_exists():
            return "❌ Сессия не найдена"
        
        try:
            size = os.path.getsize(self.session_file)
            if size < 100:
                return "⚠️ Поврежденная сессия"
            
            mtime = os.path.getmtime(self.session_file)
            date_str = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M")
            return f"✅ Сессия от {date_str}"
            
        except Exception as e:
            return f"❌ Ошибка проверки: {e}"

class TelegramAuthWindow:
    """Окно авторизации в Telegram"""
    
    def __init__(self, parent, config, theme_manager):
        self.parent = parent
        self.config = config
        self.theme_manager = theme_manager
        self.window = None
        self.result = False
        self.loop = None
        self.auth_thread = None
        
    def show_auth(self):
        """Показать окно авторизации"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("🔐 Авторизация Telegram")
        self.window.geometry("450x350")
        self.window.resizable(False, False)
        
        theme = self.theme_manager.get_current_theme()
        self.window.configure(bg=theme['bg'])
        
        self.center_window()
        self.window.transient(self.parent)
        self.window.grab_set()
        self.create_auth_interface()
        self.window.wait_window()
        return self.result
    
    def center_window(self):
        """Центрирование окна"""
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - 450) // 2
        y = (screen_height - 350) // 2
        self.window.geometry(f"450x350+{x}+{y}")
    
    def create_auth_interface(self):
        """Создание интерфейса авторизации"""
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        ttk.Label(main_frame, text="📱 Авторизация в Telegram", 
                 font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        instruction_text = """Для работы с Telegram необходима авторизация.

1. Получите API ID и Hash на my.telegram.org
2. Создайте приложение в разделе "API development tools"
3. Введите данные ниже и нажмите "Авторизоваться"
4. Введите код подтверждения из Telegram"""
        
        ttk.Label(main_frame, text=instruction_text, 
                 font=('Arial', 10), justify='left').pack(pady=(0, 20))
        
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(fields_frame, text="API ID:").grid(row=0, column=0, sticky='w', pady=5)
        self.api_id_var = tk.StringVar(value=str(self.config['telegram']['api_id']))
        ttk.Entry(fields_frame, textvariable=self.api_id_var, width=40).grid(row=0, column=1, padx=(10, 0), pady=5)
        
        ttk.Label(fields_frame, text="API Hash:").grid(row=1, column=0, sticky='w', pady=5)
        self.api_hash_var = tk.StringVar(value=self.config['telegram']['api_hash'])
        ttk.Entry(fields_frame, textvariable=self.api_hash_var, width=40).grid(row=1, column=1, padx=(10, 0), pady=5)
        
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=20)
        
        ttk.Button(buttons_frame, text="🔐 Авторизоваться", 
                  command=self.start_auth).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="❌ Отмена", 
                  command=self.cancel_auth).pack(side='right', padx=5)
        
        self.status_label = ttk.Label(main_frame, text="Готов к авторизации", 
                                     font=('Arial', 10), foreground='blue')
        self.status_label.pack(pady=10)
    
    def start_auth(self):
        """Запуск процесса авторизации"""
        try:
            api_id = int(self.api_id_var.get().strip())
            api_hash = self.api_hash_var.get().strip()
            
            if not api_id or not api_hash:
                messagebox.showerror("Ошибка", "Заполните все поля", parent=self.window)
                return

            phone_number = simpledialog.askstring("Номер телефона",
                                             "Введите номер телефона (с кодом страны):\nПример: +71234567890",
                                             parent=self.window)
            
            if not phone_number:
                self.status_label.config(text="❌ Авторизация отменена", foreground='red')
                return

            self.config['telegram']['api_id'] = api_id
            self.config['telegram']['api_hash'] = api_hash
            
            self.status_label.config(text="Запуск авторизации...", foreground='orange')
            self.window.update()
            
            self.auth_thread = threading.Thread(target=self.run_auth_loop, args=(phone_number,))
            self.auth_thread.daemon = True
            self.auth_thread.start()
            
        except ValueError:
            messagebox.showerror("Ошибка", "API ID должен быть числом", parent=self.window)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка авторизации: {e}", parent=self.window)

    async def ask_string_async(self, title, prompt, show=None):
        """Асинхронный вызов simpledialog"""
        future = self.loop.create_future()
        
        def ask():
            result = simpledialog.askstring(title, prompt, parent=self.window, show=show)
            self.loop.call_soon_threadsafe(future.set_result, result)

        self.window.after(0, ask)
        return await future

    def run_auth_loop(self, phone_number):
        """Запуск цикла авторизации"""
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_until_complete(self.perform_auth(phone_number))
        except Exception as e:
            error_message = f"Критическая ошибка: {e}"
            self.window.after(0, self.status_label.config, {'text': error_message, 'foreground': 'red'})
        finally:
            self.loop.close()

    async def perform_auth(self, phone_number):
        """Процесс авторизации"""
        app = None
        try:
            app = Client(
                self.config['telegram']['session_name'],
                api_id=self.config['telegram']['api_id'],
                api_hash=self.config['telegram']['api_hash']
            )

            self.window.after(0, self.status_label.config,
                            {'text': "Подключение...", 'foreground': 'orange'})
            await app.connect()

            # Отправляем код
            self.window.after(0, self.status_label.config,
                            {'text': "Отправка кода...", 'foreground': 'orange'})
            sent_code = await app.send_code(phone_number)

            # Запрашиваем код
            code = await self.ask_string_async(
                "Код подтверждения",
                "Введите код из Telegram или SMS:"
            )
            if not code:
                self.window.after(0, self.status_label.config,
                                {'text': "❌ Авторизация отменена", 'foreground': 'red'})
                return

            self.window.after(0, self.status_label.config,
                {'text': "Проверка кода...", 'foreground': 'orange'})
            
            try:
                await app.sign_in(phone_number, sent_code.phone_code_hash, code.strip().replace(" ", ""))
                
            except SessionPasswordNeeded:
                # Запрашиваем пароль 2FA
                self.window.after(0, self.status_label.config,
                                {'text': "Требуется пароль 2FA...", 'foreground': 'blue'})
                
                password = await self.ask_string_async(
                    "Пароль 2FA",
                    "Введите пароль двухфакторной аутентификации:",
                    show="*"
                )
                if not password:
                    self.window.after(0, self.status_label.config,
                                    {'text': "❌ Авторизация отменена", 'foreground': 'red'})
                    return
                
                self.window.after(0, self.status_label.config,
                                {'text': "Проверка пароля 2FA...", 'foreground': 'orange'})
                await app.check_password(password)

            # Проверка успешной авторизации
            me = await app.get_me()
            self.window.after(0, self.status_label.config,
                            {'text': f"✅ Авторизация успешна! ({me.first_name})", 'foreground': 'green'})
            self.result = True
            self.window.after(1500, self.close_auth)

        except Exception as e:
            error_message = f"❌ Ошибка: {e}"
            self.window.after(0, self.status_label.config,
                            {'text': error_message, 'foreground': 'red'})
            self.result = False
        finally:
            if app and app.is_connected:
                await app.disconnect()

    def cancel_auth(self):
        """Отмена авторизации"""
        self.result = False
        self.close_auth()

    def close_auth(self):
        """Закрытие окна авторизации"""
        if self.window:
            self.window.destroy()

class EmailSender:
    """Отправка сообщений по email"""
    
    def __init__(self, config):
        self.config = config['email']
        
    def send_email(self, to_email, subject, message, attachment_path=None):
        """Отправка email с возможным вложением"""
        try:
            if not self.config['username'] or not self.config['password']:
                logging.warning("Email настройки не заданы")
                return False
                    
            msg = MIMEMultipart()
            msg['From'] = self.config['username']
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'plain', 'utf-8'))
            
            # Добавляем вложение если есть
            if attachment_path and os.path.exists(attachment_path):
                try:
                    from email.mime.application import MIMEApplication
                    from email.mime.image import MIMEImage
                    from email.mime.audio import MIMEAudio
                    import mimetypes
                    
                    filename = os.path.basename(attachment_path)
                    
                    # Определяем тип файла
                    content_type, _ = mimetypes.guess_type(attachment_path)
                    
                    with open(attachment_path, 'rb') as f:
                        file_data = f.read()
                    
                    if content_type is None:
                        content_type = 'application/octet-stream'
                    
                    main_type, sub_type = content_type.split('/', 1)
                    
                    if main_type == 'image':
                        attachment = MIMEImage(file_data, _subtype=sub_type)
                    elif main_type == 'audio':
                        attachment = MIMEAudio(file_data, _subtype=sub_type)
                    else:
                        attachment = MIMEApplication(file_data)
                    
                    attachment.add_header('Content-Disposition', 'attachment', filename=filename)
                    msg.attach(attachment)
                    
                    logging.info(f"📎 Добавлено вложение: {filename}")
                    
                except Exception as attach_error:
                    logging.warning(f"⚠️ Ошибка добавления вложения: {attach_error}")

            smtp_server = self.config['smtp_server']
            smtp_port = self.config['smtp_port']
            use_tls = self.config.get('use_tls', True)
            
            server = None
            
            try:
                if smtp_port in [465, 587]:
                    if smtp_port == 465:
                        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
                    else:
                        server = smtplib.SMTP(smtp_server, smtp_port)
                        server.starttls()
                else:
                    if use_tls:
                        try:
                            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
                        except:
                            server = smtplib.SMTP(smtp_server, smtp_port)
                            server.starttls()
                    else:
                        server = smtplib.SMTP(smtp_server, smtp_port)
                
                # Аутентификация и отправка
                server.login(self.config['username'], self.config['password'])
                server.sendmail(self.config['username'], to_email, msg.as_string())
                server.quit()
                
                if attachment_path:
                    logging.info(f"Email с вложением отправлен на {to_email}")
                else:
                    logging.info(f"Email отправлен на {to_email}")
                return True
                
            except Exception as conn_error:
                if server:
                    try:
                        server.quit()
                    except:
                        pass
                raise conn_error

        except Exception as e:
            error_msg = str(e)
            if "SSL: WRONG_VERSION_NUMBER" in error_msg:
                logging.error(f"SSL ошибка для {to_email}: Попробуйте порт 587")
            elif "Authentication" in error_msg or "auth" in error_msg.lower():
                logging.error(f"Ошибка аутентификации для {to_email}")
            elif "Connection" in error_msg or "connect" in error_msg.lower():
                logging.error(f"Ошибка подключения для {to_email}")
            else:
                logging.error(f"Ошибка отправки email на {to_email}: {e}")
            return False

class UniversalSenderGUI:
    """Графический интерфейс приложения"""
    
    def __init__(self):
        self.config = UniversalSenderConfig()
        self.root = tk.Tk()
        self.root.title("Универсальная рассылка v3.1")
        self.root.geometry("700x1100")
        self.root.resizable(True, True)
        
        # Инициализация менеджера тем
        self.theme_manager = ThemeManager(self.root)
        
        # Данные для рассылки
        self.contacts = []
        self.data_df = None
        
        # Счетчики
        self.sent_count = 0
        self.total_count = 0
        
        # Переменная для остановки рассылки
        self.stop_flag = False
        # Вложения
        self.attachment_file = None
        self.attachment_filename = None
        
        self.create_widgets()
        self.theme_manager.apply_theme()
    
    def create_widgets(self):
        """Создание элементов интерфейса"""
        
        # === ЗАГОЛОВОК ===
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill='x', padx=10, pady=(10, 0))
        
        title_label = ttk.Label(header_frame, text="📡 Универсальная рассылка v3.1", 
                               font=('Arial', 14, 'bold'))
        title_label.pack(side='left')
        
        self.theme_button = ttk.Button(header_frame, text="🌙 Темная тема", 
                                      command=self.toggle_theme)
        self.theme_button.pack(side='right')
        
        # === ФАЙЛЫ ===
        files_frame = ttk.LabelFrame(self.root, text="📁 Файлы данных", padding=10)
        files_frame.pack(fill='x', padx=10, pady=5)
        
        # Первая строка - выбор и загрузка данных в одну строку
        data_row = ttk.Frame(files_frame)
        data_row.pack(fill='x', pady=(0, 10))
        
        ttk.Button(data_row, text="📊 Выбрать Excel файл", 
                  command=self.select_data_file, width=20).pack(side='left', padx=(0, 10))
        ttk.Button(data_row, text="📥 Загрузить контакты", 
                  command=self.load_contacts, width=20).pack(side='left', padx=(0, 10))
        
        self.data_label = ttk.Label(data_row, text="Файл не выбран", foreground="gray")
        self.data_label.pack(side='left', padx=(10, 0))
        
        # Информация о контактах
        self.contacts_info_label = ttk.Label(files_frame, text="", foreground="blue")
        self.contacts_info_label.pack(pady=5, anchor='w')

        
        # === НАСТРОЙКИ ОТПРАВКИ ===
        settings_frame = ttk.LabelFrame(self.root, text="⚙️ Настройки отправки", padding=10)
        settings_frame.pack(fill='x', padx=10, pady=5)
        
        # Способы отправки
        methods_frame = ttk.Frame(settings_frame)
        methods_frame.pack(fill='x', pady=5)
        
        self.send_telegram_var = tk.BooleanVar(value=self.config.config['sending']['send_telegram'])
        self.send_email_var = tk.BooleanVar(value=self.config.config['sending']['send_email'])
        
        ttk.Checkbutton(methods_frame, text="📱 Отправлять в Telegram", 
                       variable=self.send_telegram_var).pack(side='left', padx=10)
        ttk.Checkbutton(methods_frame, text="📧 Отправлять на Email", 
                       variable=self.send_email_var).pack(side='left', padx=10)
        
        # === УПРАВЛЯЮЩИЕ КНОПКИ (НОВОЕ РАСПОЛОЖЕНИЕ) ===
        control_buttons_frame = ttk.Frame(settings_frame)
        control_buttons_frame.pack(fill='x', pady=10)
        
        # Левая колонка - основные кнопки
        left_column = ttk.Frame(control_buttons_frame)
        left_column.pack(side='left', fill='y')
        
        # Первый ряд левой колонки
        left_row1 = ttk.Frame(left_column)
        left_row1.pack(fill='x', pady=(0, 5))
        
        ttk.Button(left_row1, text="🔑 TG Авторизация", 
                  command=self.show_telegram_auth, width=20).pack(side='left', padx=(0, 10))
        ttk.Button(left_row1, text="💾 Сохранить", 
                  command=self.apply_all_settings, width=20).pack(side='left')
        
        # Второй ряд левой колонки
        left_row2 = ttk.Frame(left_column)
        left_row2.pack(fill='x')
        
        ttk.Button(left_row2, text="🗑️ Очистить TG", 
                  command=self.clear_telegram_session, width=18).pack(side='left', padx=(0, 10))
        
        # Правая колонка - тест и статус
        right_column = ttk.Frame(control_buttons_frame)
        right_column.pack(side='right', fill='y', padx=(20, 0))
        
        ttk.Button(right_column, text="🧪 Тест Email", 
                  command=self.test_email_connection, width=20).pack(pady=(0, 5))
        ttk.Button(right_column, text="🔄 Обновить статус", 
                  command=self.update_telegram_status, width=20).pack()
        
        # === СТАТУС TELEGRAM ===
        tg_status_frame = ttk.Frame(settings_frame)
        tg_status_frame.pack(fill='x', pady=5)
        
        ttk.Label(tg_status_frame, text="📱 Статус Telegram:", 
                 font=('Arial', 10, 'bold')).pack(side='left')
        
        self.telegram_status_label = ttk.Label(tg_status_frame, text="❓ Проверка...", 
                                              foreground='orange')
        self.telegram_status_label.pack(side='left', padx=(10, 0))
        
        # === ЗАДЕРЖКА ===
        delay_frame = ttk.Frame(settings_frame)
        delay_frame.pack(fill='x', pady=10)
        
        ttk.Label(delay_frame, text="⏱️ Задержка между отправками:").pack(side='left')
        
        self.delay_var = tk.DoubleVar(value=self.config.config['sending']['delay_seconds'])
        self.delay_scale = ttk.Scale(delay_frame, from_=0, to=3, variable=self.delay_var, 
                                   orient='horizontal', length=200)
        self.delay_scale.pack(side='left', padx=10)
        
        self.delay_label = ttk.Label(delay_frame, text="1.0 сек")
        self.delay_label.pack(side='left')
        self.delay_var.trace('w', self.update_delay_label)
        
        # === EMAIL НАСТРОЙКИ ===
        email_frame = ttk.LabelFrame(settings_frame, text="📧 Настройки Email")
        email_frame.pack(fill='x', pady=10)
        
        # Провайдеры
        provider_frame = ttk.Frame(email_frame)
        provider_frame.pack(fill='x', pady=5)
        
        ttk.Label(provider_frame, text="Провайдер:").pack(side='left')
        
        self.email_provider_var = tk.StringVar(value=self.config.config['email']['provider'])
        
        providers = [("Gmail", "gmail"), ("Yandex", "yandex"), ("Mail.ru", "mailru"), 
                    ("Zoho", "zoho"), ("Свой SMTP", "custom")]
        
        for text, value in providers:
            ttk.Radiobutton(provider_frame, text=text, variable=self.email_provider_var, 
                           value=value, command=self.on_provider_change).pack(side='left', padx=5)
        
        # Учетные данные
        creds_frame = ttk.Frame(email_frame)
        creds_frame.pack(fill='x', pady=5)
        
        ttk.Label(creds_frame, text="Email:").grid(row=0, column=0, sticky='w', padx=5)
        self.email_username_var = tk.StringVar(value=self.config.config['email']['username'])
        ttk.Entry(creds_frame, textvariable=self.email_username_var, width=30).grid(row=0, column=1, padx=5)
        
        ttk.Label(creds_frame, text="Пароль:").grid(row=1, column=0, sticky='w', padx=5)
        self.email_password_var = tk.StringVar(value=self.config.config['email']['password'])
        ttk.Entry(creds_frame, textvariable=self.email_password_var, show='*', width=30).grid(row=1, column=1, padx=5)
        
        # Кастомные SMTP настройки
        self.custom_smtp_frame = ttk.Frame(email_frame)
        self.custom_smtp_frame.pack(fill='x', pady=5)
        
        ttk.Label(self.custom_smtp_frame, text="SMTP сервер:").grid(row=0, column=0, sticky='w', padx=5)
        self.smtp_server_var = tk.StringVar(value=self.config.config['email']['smtp_server'])
        ttk.Entry(self.custom_smtp_frame, textvariable=self.smtp_server_var, width=20).grid(row=0, column=1, padx=5)
        
        ttk.Label(self.custom_smtp_frame, text="Порт:").grid(row=0, column=2, sticky='w', padx=5)
        self.smtp_port_var = tk.IntVar(value=self.config.config['email']['smtp_port'])
        ttk.Entry(self.custom_smtp_frame, textvariable=self.smtp_port_var, width=10).grid(row=0, column=3, padx=5)
        
        self.use_tls_var = tk.BooleanVar(value=self.config.config['email']['use_tls'])
        ttk.Checkbutton(self.custom_smtp_frame, text="Использовать шифрование", 
                       variable=self.use_tls_var).grid(row=1, column=0, columnspan=4, sticky='w', pady=5)
        
        self.on_provider_change()
        
        # === СООБЩЕНИЕ ===
        message_frame = ttk.LabelFrame(self.root, text="✉️ Сообщение", padding=10)
        message_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Тема
        subject_frame = ttk.Frame(message_frame)
        subject_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(subject_frame, text="📝 Тема сообщения (для email):").pack(anchor='w')
        self.subject_var = tk.StringVar(value=self.config.config['message']['subject'])
        ttk.Entry(subject_frame, textvariable=self.subject_var, font=('Arial', 10)).pack(fill='x', pady=5)
        
        # Вложения
        attachment_frame = ttk.Frame(message_frame)
        attachment_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(attachment_frame, text="📎 Вложение (до 50MB):").pack(anchor='w')
        
        attachment_control_frame = ttk.Frame(attachment_frame)
        attachment_control_frame.pack(fill='x', pady=5)
        
        ttk.Button(attachment_control_frame, text="📁 Выбрать файл", 
                  command=self.select_attachment_file, width=20).pack(side='left', padx=(0, 10))
        ttk.Button(attachment_control_frame, text="🗑️ Убрать файл", 
                  command=self.remove_attachment_file, width=20).pack(side='left')
        
        self.attachment_label = ttk.Label(attachment_control_frame, text="Файл не выбран", 
                                         foreground="gray")
        self.attachment_label.pack(side='left', padx=(10, 0))
        
        # Текст
        ttk.Label(message_frame, text="📄 Шаблон сообщения (переменные: {name}):").pack(anchor='w')
        self.message_text = scrolledtext.ScrolledText(message_frame, height=10, wrap='word',
                                                     font=('Arial', 10))
        self.message_text.pack(fill='both', expand=True, pady=5)
        self.message_text.insert('1.0', self.config.config['message']['template'])
        
        # === УПРАВЛЕНИЕ РАССЫЛКОЙ ===
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(control_frame, text="▶ Запустить рассылку", 
                  command=self.start_sending).pack(side='left', padx=5)
        ttk.Button(control_frame, text="⏹ Остановить", 
                  command=self.stop_sending).pack(side='left', padx=5)
        ttk.Button(control_frame, text="💾 Сохранить все", 
                  command=self.save_all_settings).pack(side='right', padx=5)
        ttk.Button(control_frame, text="❓ Help", 
                  command=self.show_help_window).pack(side='right', padx=5)
        
        # === СТАТУС И ЛОГИ ===
        status_frame = ttk.LabelFrame(self.root, text="📊 Статус и логи", padding=10)
        status_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        stats_frame = ttk.Frame(status_frame)
        stats_frame.pack(fill='x', pady=5)
        
        self.stats_label = ttk.Label(stats_frame, text="Готов к работе")
        self.stats_label.pack(side='left')
        
        self.progress_bar = ttk.Progressbar(stats_frame, mode='determinate')
        self.progress_bar.pack(side='right', fill='x', expand=True, padx=(10, 0))
        
        self.log_text = scrolledtext.ScrolledText(status_frame, height=12, state='disabled')
        self.log_text.pack(fill='both', expand=True, pady=5)
        
        # Переменная для остановки рассылки
        self.stop_flag = False
        
        # Обновляем статус Telegram при запуске
        self.root.after(1000, self.update_telegram_status)
    
    def toggle_theme(self):
        """Переключение темы"""
        theme_name = self.theme_manager.toggle_theme()
        self.theme_button.config(text=theme_name)
        
        # Применяем тему к ScrolledText виджетам
        theme = self.theme_manager.get_current_theme()
        try:
            if self.theme_manager.is_dark:
                self.log_text.configure(bg=theme['log_bg'], fg=theme['fg'], 
                                    insertbackground=theme['fg'])
                self.message_text.configure(bg=theme['text_bg'], fg=theme['fg'],
                                        insertbackground=theme['fg'])
            else:
                self.log_text.configure(bg=theme['text_bg'], fg=theme['fg'])
                self.message_text.configure(bg=theme['text_bg'], fg=theme['fg'])
        except:
            pass
        
        # ДОБАВИТЬ ЭТО: Принудительное обновление всех виджетов
        self.root.configure(bg=theme['bg'])
        self.root.update_idletasks()
        
        # Обновляем все дочерние фреймы принудительно
        for widget in self.root.winfo_children():
            self._update_widget_theme(widget, theme)
            
    def _update_widget_theme(self, widget, theme):
        """Рекурсивное обновление темы для всех виджетов"""
        try:
            # Обновляем обычные tk виджеты
            if isinstance(widget, (tk.Frame, tk.Label)):
                widget.configure(bg=theme['bg'])
                if isinstance(widget, tk.Label):
                    widget.configure(fg=theme['fg'])
            
            # Рекурсивно обновляем дочерние виджеты
            for child in widget.winfo_children():
                self._update_widget_theme(child, theme)
                
        except Exception:
            pass 
    
    def apply_all_settings(self):
        """Применение всех настроек"""
        # Email настройки
        self.config.config['email']['provider'] = self.email_provider_var.get()
        self.config.config['email']['username'] = self.email_username_var.get()
        self.config.config['email']['password'] = self.email_password_var.get()
        self.config.config['email']['smtp_server'] = self.smtp_server_var.get()
        self.config.config['email']['smtp_port'] = self.smtp_port_var.get()
        self.config.config['email']['use_tls'] = self.use_tls_var.get()
        
        # Настройки отправки
        self.config.config['sending']['send_telegram'] = self.send_telegram_var.get()
        self.config.config['sending']['send_email'] = self.send_email_var.get()
        self.config.config['sending']['delay_seconds'] = self.delay_var.get()
        
        self.log("✅ Настройки применены")
    
    def test_email_connection(self):
        """Тестирование email подключения"""
        try:
            self.apply_all_settings()
            
            smtp_server = self.config.config['email']['smtp_server']
            smtp_port = self.config.config['email']['smtp_port']
            username = self.config.config['email']['username']
            password = self.config.config['email']['password']
            use_tls = self.config.config['email'].get('use_tls', True)
            
            if not username or not password:
                messagebox.showerror("Ошибка", "Введите email и пароль")
                return
            
            self.log("🧪 Тестирование email подключения...")
            
            server = None
            if smtp_port == 465:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port)
                if use_tls:
                    server.starttls()
            
            server.login(username, password)
            server.quit()
            
            self.log("🎉 Email тест успешен!")
            messagebox.showinfo("Успех", "Подключение к email серверу работает!")
                
        except Exception as e:
            error_msg = f"Ошибка email теста: {e}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("Ошибка", error_msg)
    
    def show_telegram_auth(self):
        """Показать окно авторизации Telegram"""
        auth_window = TelegramAuthWindow(self.root, self.config.config, self.theme_manager)
        if auth_window.show_auth():
            self.log("✅ Авторизация Telegram успешна!")
            self.config.save_config()
            self.update_telegram_status()
        else:
            self.log("❌ Авторизация Telegram отменена")
    
    def update_telegram_status(self):
        """Обновление статуса Telegram сессии"""
        try:
            temp_sender = TelegramSender(self.config.config)
            status = temp_sender.get_session_status()
            
            # Определяем цвет статуса
            if "✅" in status:
                color = 'green'
            elif "⚠️" in status:
                color = 'orange'
            elif "❌" in status:
                color = 'red'
            else:
                color = 'blue'
            
            self.telegram_status_label.config(text=status, foreground=color)
            
        except Exception as e:
            error_status = f"❌ Ошибка проверки: {e}"
            self.telegram_status_label.config(text=error_status, foreground='red')

    def clear_telegram_session(self):
        """Очистка Telegram сессии"""
        try:
            session_file = f"{self.config.config['telegram']['session_name']}.session"
            if os.path.exists(session_file):
                os.remove(session_file)
                self.log("🗑️ Файл сессии Telegram удален")
                self.update_telegram_status()
                messagebox.showinfo("Успех", "Сессия очищена. Потребуется повторная авторизация.")
            else:
                messagebox.showinfo("Информация", "Файл сессии не найден")
        except Exception as e:
            error_msg = f"Ошибка удаления сессии: {e}"
            self.log(f"❌ {error_msg}")
            messagebox.showerror("Ошибка", error_msg)
    
    def update_delay_label(self, *args):
        """Обновление метки задержки"""
        delay = self.delay_var.get()
        self.delay_label.config(text=f"{delay:.1f} сек")
    
    def on_provider_change(self):
        """Обработка изменения провайдера"""
        provider = self.email_provider_var.get()
        
        if provider == 'custom':
            self.custom_smtp_frame.pack(fill='x', pady=5)
        else:
            self.custom_smtp_frame.pack_forget()
            # Применяем настройки выбранного провайдера
            if provider in self.config.config['smtp_presets']:
                preset = self.config.config['smtp_presets'][provider]
                self.smtp_server_var.set(preset['smtp_server'])
                self.smtp_port_var.set(preset['smtp_port'])
                self.use_tls_var.set(preset['use_tls'])
    
    def select_data_file(self):
        """Выбор файла данных"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл с контактами",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.data_file = filename
            self.data_label.config(text=os.path.basename(filename))
    
    def load_contacts(self):
        """Загрузка контактов из файла"""
        try:
            if not hasattr(self, 'data_file'):
                messagebox.showerror("Ошибка", "Сначала выберите файл")
                return
            
            self.log("🔄 Загрузка контактов...")
            
            self.data_df = DataProcessor.load_excel_file(self.data_file)
            self.contacts = DataProcessor.process_contacts(self.data_df)
            
            self.total_count = len(self.contacts)
            self.sent_count = 0
            self.update_stats()
            
            # Показываем информацию о найденных контактах
            telegram_count = sum(1 for c in self.contacts if c['telegram'])
            email_count = sum(1 for c in self.contacts if c['email'])
            
            info_text = f"📊 Загружено: {len(self.contacts)} | 📱 TG: {telegram_count} | 📧 Email: {email_count}"
            self.contacts_info_label.config(text=info_text)
            
            self.log(f"✅ Контакты загружены: {len(self.contacts)} записей")
            
        except Exception as e:
            error_msg = f"Ошибка загрузки контактов: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("Ошибка", error_msg)
    
    def save_all_settings(self):
        """Сохранение всех настроек"""
        self.apply_all_settings()
        
        # Сохраняем сообщение
        self.config.config['message']['subject'] = self.subject_var.get()
        self.config.config['message']['template'] = self.message_text.get('1.0', 'end-1c')
        
        self.config.save_config()
        self.log("💾 Все настройки сохранены")
        messagebox.showinfo("Успех", "Все настройки сохранены")
    
    def log(self, message):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        
        self.log_text.config(state='normal')
        self.log_text.insert('end', log_message + '\n')
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        
        logging.info(message)
        self.root.update()
    
    def update_stats(self):
        """Обновление статистики"""
        stats_text = f"📊 Контакты: {self.total_count} | ✅ Отправлено: {self.sent_count}"
        self.stats_label.config(text=stats_text)
        
        if self.total_count > 0:
            progress = (self.sent_count / self.total_count) * 100
            self.progress_bar['value'] = progress
    
    def select_attachment_file(self):
        """Выбор файла для вложения"""
        filename = filedialog.askopenfilename(
            title="Выберите файл для вложения",
            filetypes=[
                ("Все файлы", "*.*"),
                ("Изображения", "*.jpg *.jpeg *.png *.gif *.bmp"),
                ("Документы", "*.pdf *.doc *.docx *.txt *.rtf"),
                ("Архивы", "*.zip *.rar *.7z"),
                ("Видео", "*.mp4 *.avi *.mkv *.mov"),
                ("Аудио", "*.mp3 *.wav *.flac *.ogg")
            ]
        )
        
        if filename:
            # Проверка размера файла
            try:
                file_size = os.path.getsize(filename)
                max_size = 50 * 1024 * 1024  # 50MB в байтах
                
                if file_size > max_size:
                    size_mb = file_size / (1024 * 1024)
                    messagebox.showerror("Ошибка", 
                                       f"Файл слишком большой: {size_mb:.1f}MB\n"
                                       f"Максимальный размер: 50MB")
                    return
                
                self.attachment_file = filename
                self.attachment_filename = os.path.basename(filename)
                
                # Показываем информацию о файле
                size_mb = file_size / (1024 * 1024)
                if size_mb < 1:
                    size_text = f"{file_size / 1024:.1f}KB"
                else:
                    size_text = f"{size_mb:.1f}MB"
                
                display_text = f"{self.attachment_filename} ({size_text})"
                self.attachment_label.config(text=display_text, foreground="green")
                
                self.log(f"📎 Выбран файл: {self.attachment_filename} ({size_text})")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка при проверке файла: {e}")
    
    def remove_attachment_file(self):
        """Удаление выбранного вложения"""
        if self.attachment_file:
            self.attachment_file = None
            self.attachment_filename = None
            self.attachment_label.config(text="Файл не выбран", foreground="gray")
            self.log("🗑️ Вложение удалено")
        else:
            messagebox.showinfo("Информация", "Файл не был выбран")
    def start_sending(self):
        """Запуск рассылки"""
        if not self.contacts:
            messagebox.showerror("Ошибка", "Сначала загрузите контакты")
            return
        
        if not self.send_telegram_var.get() and not self.send_email_var.get():
            messagebox.showerror("Ошибка", "Выберите хотя бы один способ отправки")
            return
        
        # Применяем настройки перед запуском
        self.apply_all_settings()
        
        self.stop_flag = False
        self.sent_count = 0
        self.update_stats()
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.run_sending)
        thread.daemon = True
        thread.start()
    
    def stop_sending(self):
        """Остановка рассылки"""
        self.stop_flag = True
        self.log("⏹ Остановка рассылки...")
    
    def run_sending(self):
        """Основной процесс рассылки"""
        try:
            asyncio.run(self.send_notifications())
        except Exception as e:
            self.log(f"❌ Критическая ошибка рассылки: {e}")
    
    async def send_notifications(self):
        """Отправка уведомлений"""
        message_template = self.message_text.get('1.0', 'end-1c')
        delay_seconds = self.delay_var.get()
        subject = self.subject_var.get()
        
        # Инициализация отправителей
        telegram_sender = None
        email_sender = None
        
        if self.send_telegram_var.get():
            telegram_sender = TelegramSender(self.config.config)
            if not await telegram_sender.initialize():
                self.log("⚠️ Telegram не авторизован. Требуется авторизация.")
                self.root.after(0, self.show_auth_prompt)
                return
        
        if self.send_email_var.get():
            email_sender = EmailSender(self.config.config)
        
        attachment_info = ""
        if self.attachment_file:
            attachment_info = f" с вложением ({self.attachment_filename})"
        
        self.log(f"🚀 Начинается рассылка для {len(self.contacts)} получателей{attachment_info}")
        
        for i, contact in enumerate(self.contacts):
            if self.stop_flag:
                break
            
            try:
                # Формирование сообщения
                message = message_template.format(name=contact['name'])
                
                sent_methods = []
                
                # Отправка в Telegram
                if telegram_sender and contact['telegram']:
                    if await telegram_sender.send_message_with_attachment(
                        contact['telegram'], message, self.attachment_file):
                        sent_methods.append("TG")
                
                # Отправка по Email
                if email_sender and contact['email']:
                    personalized_subject = subject.format(name=contact['name']) if '{name}' in subject else subject
                    if email_sender.send_email(contact['email'], personalized_subject, 
                                             message, self.attachment_file):
                        sent_methods.append("Email")
                
                if sent_methods:
                    self.sent_count += 1
                    methods_text = "/".join(sent_methods)
                    self.log(f"✅ {contact['name']} ({methods_text})")
                else:
                    self.log(f"❌ {contact['name']} - не удалось отправить")
                
                self.update_stats()
                
                # Задержка между отправками
                if delay_seconds > 0:
                    await asyncio.sleep(delay_seconds)
                
            except Exception as e:
                self.log(f"❌ Ошибка для {contact['name']}: {e}")
        
        # Закрытие соединений
        if telegram_sender:
            await telegram_sender.close()
        
        if self.stop_flag:
            self.log("⏹ Рассылка остановлена пользователем")
        else:
            self.log(f"🎉 Рассылка завершена! Отправлено: {self.sent_count}/{self.total_count}")
    
    def show_auth_prompt(self):
        """Показать запрос на авторизацию Telegram"""
        response = messagebox.askyesno("Авторизация Telegram", 
                                      "Telegram не авторизован. Пройти авторизацию сейчас?")
        if response:
            self.show_telegram_auth()
    
    def show_help_window(self):
        """Показать окно помощи"""
        if not hasattr(self, 'help_window'):
            self.help_window = HelpWindow(self.root, self.theme_manager)
        self.help_window.show_help()
    
    def run(self):
        """Запуск приложения"""
        self.root.mainloop()

class HelpWindow:
    """Окно помощи с анимированным email автора"""
    
    def __init__(self, parent, theme_manager):
        self.parent = parent
        self.theme_manager = theme_manager
        self.window = None
        self.animation_running = False
        self.original_email = "fxranxklin@proton.me"
        
    def show_help(self):
        """Показать окно помощи"""
        if self.window and self.window.winfo_exists():
            self.window.lift()
            return
            
        self.window = tk.Toplevel(self.parent)
        self.window.title("📖 Справка - Универсальная рассылка v3.1")
        self.window.geometry("650x700")
        self.window.resizable(True, True)
        
        # Применяем текущую тему
        theme = self.theme_manager.get_current_theme()
        self.window.configure(bg=theme['bg'])
        self.window.resizable(False, False)
        
        # Центрирование окна
        self.center_window()
        
        # Создание содержимого
        self.create_content()
        
        # Запуск анимации
        self.start_animation()
        
        # Обработка закрытия окна
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def center_window(self):
        """Центрирование окна относительно родительского"""
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # Получаем размеры экрана и окна
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        x = (screen_width - 60) // 2
        y = (screen_height - 700) // 2
        
        self.window.geometry(f"540x970+{x}+{y}")

    def create_content(self):
        """Создание содержимого окна помощи"""
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill='x', pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="📡 Универсальная рассылка", 
                               font=('Arial', 16, 'bold'))
        title_label.pack()
        
        version_label = ttk.Label(title_frame, text="Версия 3.1 FIXED", 
                                 font=('Arial', 10), foreground='gray')
        version_label.pack()
        
        # Основная информация
        info_text = """🚀 Универсальная программа для массовых рассылок
📱 Одновременная отправка в Telegram и Email
📊 Автоматический парсинг Excel файлов любой сложности
📧 Поддержка всех популярных email провайдеров
🌙 Темная и светлая темы интерфейса

🔧 ОСНОВНЫЕ ФУНКЦИИ:
• Автопоиск колонок: Username, FIO, Email
• Проверка статуса Telegram сессии
• Тестирование email подключения
• Регулируемая задержка между отправками
• Подробные логи всех операций

⚙️ БЫСТРЫЙ СТАРТ:
1. Выберите Excel файл и загрузите контакты
2. Для Telegram: нажмите "🔑 TG Авторизация"
3. Для Email: настройте провайдера и нажмите "🧪 Тест Email"
4. Нажмите "💾 Применить настройки"
5. Запустите рассылку кнопкой "▶ Запустить"

📞 ПОДДЕРЖКА:
Если возникли проблемы - проверьте логи внизу программы.
Программа автоматически диагностирует большинство ошибок."""
        
        info_label = ttk.Label(main_frame, text=info_text, 
                              font=('Arial', 10), justify='left')
        info_label.pack(pady=(0, 20))
        
        # Анимированный email автора
        author_frame = ttk.Frame(main_frame)
        author_frame.pack(pady=20)
        
        ttk.Label(author_frame, text="👨‍💻 Автор:", font=('Arial', 12, 'bold')).pack()
        
        email_frame = ttk.Frame(author_frame)
        email_frame.pack(pady=10)
        
        theme = self.theme_manager.get_current_theme()
        self.email_canvas = tk.Canvas(email_frame, width=400, height=50, 
                                    bg=theme['bg'], highlightthickness=0)
        self.email_canvas.pack()
        
        self.init_letter_positions()
        
        # Секция донатов
        self.create_donation_section(main_frame)
        
        # Кнопка закрытия
        ttk.Button(main_frame, text="✅ Понятно", command=self.on_close).pack(pady=20)
        
    def create_donation_section(self, parent_frame):
        """Создание секции донатов"""
        donation_frame = ttk.LabelFrame(parent_frame, text="💰 Поддержать разработчика", padding=15)
        donation_frame.pack(fill='x', pady=(20, 10))
        
        ttk.Label(donation_frame, 
                 text="☕ Если программа оказалась полезной, можете угостить кофе:", 
                 font=('Arial', 10), justify='center').pack(pady=(0, 15))
        
        # USDT адрес
        usdt_frame = ttk.Frame(donation_frame)
        usdt_frame.pack(fill='x', pady=8)
        
        ttk.Label(usdt_frame, text="🟢 USDT (TRC20):", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        usdt_address = "TSE6qB5efLq3bVuTrnU3FRvtEYQ9ktjUj4"
        theme = self.theme_manager.get_current_theme()
        usdt_label = tk.Label(usdt_frame, text=usdt_address, 
                             font=('Consolas', 10, 'underline'), 
                             fg='blue', cursor='hand2', bg=theme['bg'])
        usdt_label.pack(anchor='w', pady=(2, 0))
        
        def copy_usdt(event=None):
            self.window.clipboard_clear()
            self.window.clipboard_append(usdt_address)
            original_text = usdt_label.cget('text')
            usdt_label.config(text="✅ Скопировано!", fg='green')
            self.window.after(2000, lambda: usdt_label.config(text=original_text, fg='blue'))
        
        usdt_label.bind("<Button-1>", copy_usdt)
        
        # BTC адрес
        btc_frame = ttk.Frame(donation_frame)
        btc_frame.pack(fill='x', pady=8)
        
        ttk.Label(btc_frame, text="🟠 Bitcoin (BTC):", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        btc_address = "1HS2bhG8VfMYaswFR3ww1yuXc9kb3hPteJ"
        btc_label = tk.Label(btc_frame, text=btc_address, 
                            font=('Consolas', 10, 'underline'), 
                            fg='blue', cursor='hand2', bg=theme['bg'])
        btc_label.pack(anchor='w', pady=(2, 0))
        
        def copy_btc(event=None):
            self.window.clipboard_clear()
            self.window.clipboard_append(btc_address)
            original_text = btc_label.cget('text')
            btc_label.config(text="✅ Скопировано!", fg='green')
            self.window.after(2000, lambda: btc_label.config(text=original_text, fg='blue'))
        
        btc_label.bind("<Button-1>", copy_btc)
        
        ttk.Label(donation_frame, text="👆 Кликните по адресу чтобы скопировать", 
                 font=('Arial', 9, 'italic'), foreground='gray').pack(pady=(10, 5))
        
    def init_letter_positions(self):
        """Инициализация позиций букв email"""
        self.letter_positions = []
        email_length = len(self.original_email)
        canvas_width = 400
        
        total_width = email_length * 15
        start_x = (canvas_width - total_width) // 2
        
        for i, char in enumerate(self.original_email):
            pos = {
                'char': char,
                'x': start_x + i * 15,
                'y': 25,
                'index': i
            }
            self.letter_positions.append(pos)
        
        self.animation_time = 0

    def get_rainbow_color(self, position, time_offset=0):
        """Получить радужный цвет для позиции"""
        hue = (position * 0.1 + time_offset) % 1.0
        r, g, b = colorsys.hsv_to_rgb(hue, 0.8, 0.9)
        return f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"
    
    def start_animation(self):
        """Запуск анимации email"""
        if not self.animation_running:
            self.animation_running = True
            self.animate_letters()
    
    def animate_letters(self):
        """Анимация радужных цветов букв email"""
        if not self.animation_running or not self.window or not self.window.winfo_exists():
            return
        
        self.email_canvas.delete("all")
        self.animation_time += 0.05
        
        for i, letter in enumerate(self.letter_positions):
            color = self.get_rainbow_color(i, self.animation_time)
            y_offset = math.sin(self.animation_time * 2 + i * 0.3) * 2
            
            self.email_canvas.create_text(
                letter['x'], 
                letter['y'] + y_offset, 
                text=letter['char'], 
                font=('Consolas', 16, 'bold'), 
                fill=color, 
                anchor='center'
            )
        
        self.window.after(50, self.animate_letters)
    
    def on_close(self):
        """Закрытие окна помощи"""
        self.animation_running = False
        if self.window:
            self.window.destroy()
            self.window = None

def main():
    """Главная функция"""
    print("📡 Универсальная рассылка v3.1 FIXED")
    print("=" * 40)
    print("🔧 Исправления:")
    print("   • Убраны дублирующие кнопки")
    print("   • Исправлена работа с Telegram сессией")
    print("   • Улучшен парсинг Excel файлов")
    print("   • Переработан интерфейс")
    print("🖥️ Запуск GUI...")
    
    app = UniversalSenderGUI()
    app.run()

if __name__ == "__main__":
    main()