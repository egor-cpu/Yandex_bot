import yadisk
from aiogram import Bot, Dispatcher, types
from dotenv import load_dotenv
from aiogram import types
from aiogram.filters import Command
import os

from aiogram.types import ReplyKeyboardRemove, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton

load_dotenv("passwords.env")
BOT_TOKEN = os.getenv('BOT_TOKEN')
YANDEX_TOKEN = os.getenv('YANDEX_TOKEN')
bot = Bot(BOT_TOKEN)
dp = Dispatcher()
y = yadisk.YaDisk(token=YANDEX_TOKEN)

import logging
log_file_path = 'logs/logs_bot.log'  # Указываем имя файла для логов

# Создание обработчика для записи логов в файл на уровне WARNING и выше
file_handler = logging.FileHandler(log_file_path)
file_handler.setLevel(logging.WARNING)

# Создание обработчика для вывода логов в консоль на уровне INFO и выше
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Форматирование логов
formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Добавляем обработчики к логгеру
logging.basicConfig(level=logging.DEBUG, handlers=[file_handler, console_handler])

from create_report import *
from download_report import *
from edit_report import *
from update_temple import *
from Registration import *
button_edit_report = KeyboardButton(text="Изменить Отчёт")
button_download_report = KeyboardButton(text="Выгрузить отчёт")
button_create_report = KeyboardButton(text="Создать отчёт")
button_update_temple = KeyboardButton(text="Обновить шаблоны")
button_registration = KeyboardButton(text="Зарегистрироваться")

greet_kb = ReplyKeyboardMarkup(keyboard=[[button_edit_report], [button_download_report], [button_create_report], [button_update_temple], [button_registration]], resize_keyboard=True)

@dp.message(Command(commands=["start"]))
async def start_command_handler(message: types.Message):
    await message.answer("Привет!", reply_markup=greet_kb)

if __name__ == '__main__':
    dp.run_polling(bot)
