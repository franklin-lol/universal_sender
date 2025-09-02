import asyncio
import logging
import os
from datetime import datetime
from getpass import getpass

from pyrogram import Client
from pyrogram.errors import (
    FloodWait,
    PasswordHashInvalid,
    PhoneCodeExpired,
    PhoneCodeInvalid,
    PhoneNumberInvalid,
    SessionPasswordNeeded,
    BadRequest,
)

LOG_FILE = "session_creator.log"

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler()
        ],
    )
    logging.info("=== Pyrogram Session Creator ===")

async def main():
    setup_logging()

    default_session = f"pyro_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    session_name = input(f"Имя сессии [{default_session}]: ").strip() or default_session

    api_id = int(input("API ID: ").strip())
    api_hash = input("API HASH: ").strip()
    phone = input("Телефон (с +, напр. +71234567890): ").strip()

    logging.info(f"Создаём клиент: session='{session_name}'")
    app = Client(session_name, api_id=api_id, api_hash=api_hash)

    try:
        await app.connect()

        # --- Шаг 1: отправляем код ---
        logging.info("Отправляем код подтверждения...")
        sent_code = await app.send_code(phone)
        print("📲 Код отправлен! Проверь Telegram или SMS.")

        # --- Шаг 2: вводим код ---
        code = input("Введите код из Telegram/SMS: ").strip().replace(" ", "")
        try:
            await app.sign_in(phone, sent_code.phone_code_hash, code)
        except SessionPasswordNeeded:
            pwd = getpass("Введите пароль 2FA: ")
            try:
                await app.check_password(pwd)
            except PasswordHashInvalid:
                print("❌ Неверный пароль 2FA")
                return

        # --- Проверка ---
        me = await app.get_me()
        print(f"✅ Авторизовано как: {me.first_name} (id: {me.id})")

    except PhoneNumberInvalid:
        print("❌ Неверный номер телефона.")
    except PhoneCodeInvalid:
        print("❌ Неверный код подтверждения.")
    except PhoneCodeExpired:
        print("❌ Код подтверждения просрочен.")
    except FloodWait as e:
        print(f"⏳ Flood wait: подожди {e.value} сек.")
    except BadRequest as e:
        print(f"❌ BadRequest: {e}")
    except Exception as e:
        print(f"❌ Неизвестная ошибка: {e}")
    finally:
        try:
            await app.disconnect()
        except:
            pass

        session_file = f"{session_name}.session"
        if os.path.exists(session_file):
            print(f"📦 Файл сессии сохранён: {os.path.abspath(session_file)}")
        else:
            print("⚠️ Файл сессии не создан.")

if __name__ == "__main__":
    asyncio.run(main())
