# 📖 Подробная инструкция по установке и настройке

## 🔧 Системные требования

- **Python**: 3.8 или выше
- **ОС**: Windows 10/11, Linux, macOS
- **RAM**: минимум 512MB свободной памяти
- **Интернет**: для отправки сообщений

## 📥 Установка

### Вариант 1: Быстрая установка (рекомендуется)

```bash
# Клонирование репозитория
git clone https://github.com/yourusername/universal-sender.git
cd universal-sender

# Автоустановка зависимостей
pip install -r requirements.txt

# Запуск
python universal_sender.py
```

### Вариант 2: Ручная установка

```bash
pip install pyrogram==2.0.106
pip install pandas>=2.1.0  
pip install openpyxl>=3.1.0
pip install TgCrypto>=1.2.0
```

## 🔑 Настройка Telegram

### Получение API ключей

1. **Откройте** [my.telegram.org](https://my.telegram.org)
2. **Войдите** в свой аккаунт Telegram
3. **Перейдите** в раздел "API development tools"
4. **Создайте** новое приложение:
   - **App title**: `Universal Sender`
   - **Short name**: `universal_sender`
   - **Platform**: Desktop
5. **Скопируйте** API ID и API Hash

### Первая авторизация

**Вариант 1: Через основную программу**
1. Запустите `python universal_sender.py`
2. Нажмите "🔒 TG Авторизация"
3. Следуйте инструкциям

**Вариант 2: Отдельно создать сессию**
1. Запустите `python make_pyrogram_session.py`
2. Введите API данные и номер телефона
3. Введите код из Telegram
4. Файл сессии будет создан автоматически

## 📧 Настройка Email провайдеров

### Gmail
```
SMTP: smtp.gmail.com
Порт: 587
TLS: ✓
```

**Важно для Gmail:**
1. Включите двухфакторную аутентификацию
2. Создайте пароль приложения:
   - Google Account → Безопасность → Пароли приложений
   - Выберите "Почта" 
   - Используйте сгенерированный пароль

### Yandex
```
SMTP: smtp.yandex.ru  
Порт: 587
TLS: ✓
```

### Mail.ru
```
SMTP: smtp.mail.ru
Порт: 587  
TLS: ✓
```

### Zoho
```
SMTP: smtp.zoho.eu
Порт: 587
TLS: ✓
```

## 📊 Подготовка Excel данных

### Требования к файлу
- **Формат**: .xlsx или .xls
- **Кодировка**: UTF-8 (автоматически)
- **Размер**: неограниченно

### Названия колонок
Программа ищет колонки по ключевым словам (регистр не важен):

```
Telegram: username, tg name, tg_name, telegram, nick
Email: email, mail, почта, e-mail  
Имя: fio, имя, name, фио, tg name
```

### Примеры правильных заголовков
```
✅ Username | Email | Name
✅ TG Name | Mail | FIO  
✅ telegram | email | имя
✅ Nick | Почта | ФИО
```

### Формат данных
```
Username: @ivan_petrov (с @ или без)
Email: ivan@example.com
Name: Иван Петров
```

## 🚀 Первый запуск

### 1. Проверка установки
```bash
python --version  # Должно быть 3.8+
python -c "import pyrogram; print('OK')"
```

### 2. Запуск программы
```bash
python universal_sender.py
```

### 3. Проверка функций
1. **Загрузите тестовый Excel** с 2-3 контактами
2. **Проведите авторизацию Telegram**
3. **Протестируйте email подключение**
4. **Отправьте тестовое сообщение**

## ⚠️ Возможные проблемы

### Python не найден
```bash
# Windows
python --version
# Если ошибка, переустановите Python с python.org

# Linux/macOS
python3 --version
python3 universal_sender.py
```

### Ошибки при установке зависимостей
```bash
# Обновление pip
python -m pip install --upgrade pip

# Установка с флагами
pip install --user -r requirements.txt

# Для Linux может потребоваться
sudo apt-get install python3-tk
```

### Проблемы с Telegram
- **Сессия повреждена**: Нажмите "🗑️ Очистить TG" и авторизуйтесь заново
- **API ошибки**: Проверьте правильность API ID/Hash
- **Flood wait**: Увеличьте задержку между сообщениями

### Проблемы с Email
- **SSL ошибки**: Попробуйте порт 465 вместо 587
- **Аутентификация Gmail**: Используйте пароль приложения, не основной пароль
- **Блокировка**: Некоторые провайдеры блокируют массовые рассылки

## 📁 Структура файлов после установки

```
universal-sender/
├── universal_sender.py      # Основная программа
├── make_pyrogram_session.py # Создание сессий  
├── requirements.txt         # Зависимости
├── config.json             # Настройки (создается автоматически)
├── universal_sender.log     # Главный лог
├── session_creator.log      # Лог создания сессий
├── pyro_session_*.session   # Telegram сессии
└── README.md               # Документация
```

## 🆘 Получение помощи

1. **Проверьте логи** внизу программы
2. **Попробуйте тестовые функции** ("🧪 Тест Email", "📄 Обновить статус")
3. **Создайте Issue** в GitHub с описанием проблемы
4. **Напишите email**: fxranxklin@proton.me

---

**Удачной рассылки! 🎉**