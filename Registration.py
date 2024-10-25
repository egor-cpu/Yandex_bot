import asyncio

from aiogram import Bot, Dispatcher, types, Router, F
from aiogram.filters import Command
from aiogram.types import Message
from aiogram.types import FSInputFile
from dotenv import load_dotenv
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.filters import Command
import logging
import re
from datetime import datetime

import os
from main import bot, dp
load_dotenv("passwords.env")
Password = os.getenv('PASSWORD')

def ensure_file_exists(filename):
    if not os.path.exists(filename):
        with open(filename, 'w') as f:
            pass  # The file is now created.
class RegistrationStates(StatesGroup):
    waiting_for_password = State()
@dp.message(lambda message: message.text == "Зарегистрироваться")
async def process_reg_command(message: Message, state: FSMContext):
    ensure_file_exists('admins.txt')
    id = message.chat.id
    file = open("admins.txt", "r")
    notreg = 1
    for line in file:
        if str(id) in line:
            notreg = 0
            await bot.send_message(message.chat.id, "Вы уже зарегистрированны")
    file.close()
    if notreg == 1:
        await bot.send_message(message.chat.id, "Напишите пароль")
        await state.set_state(RegistrationStates.waiting_for_password)  # Устанавливаем состояние

@dp.message(RegistrationStates.waiting_for_password)
async def name_get(message: Message, state:FSMContext):
    if message.text == Password:
        file = open("admins.txt", "a")
        file.write(str(message.chat.id) + "\n")
        file.close()
        await message.answer("Вы успешно зарегестрировались")
        await state.clear()
    else:
        await message.answer("Извините не правильный пароль")