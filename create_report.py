import openpyxl
from main import dp, bot
import yadisk
from aiogram import Bot, Dispatcher, types
from dotenv import load_dotenv
from aiogram import types
from aiogram.filters import Command
import os
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.types import ReplyKeyboardRemove, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
import shutil
class createstate(StatesGroup):
    waiting_for_choose_template = State()
    waiting_for_choose_file_name = State()

lengh = 0
@dp.message(lambda message: message.text == "Создать отчёт")
async def create_report(message: types.Message, state: FSMContext):
    file = open("admins.txt", 'r')
    booli = 0
    for i in file:
        if i == str(message.chat.id) + "\n":
            booli = 1
    if booli == 0:
        await message.answer("У вас недостаточно доступа")
        await state.clear()
    else:
        await message.answer("Выберите шаблон\n")
        files = os.listdir("Шаблоны")
        global lengh
        a = len(files)
        lengh = a
        print(lengh)
        for i in range(a):
            await message.answer(str(i + 1) + ". " + str(files[i]))
        await state.set_state(createstate.waiting_for_choose_template)

@dp.message(createstate.waiting_for_choose_template)
async def template_get(message: types.Message, state:FSMContext):
    folder_path = 'Шаблоны'
    files = os.listdir("Шаблоны")
    file_name = message.text
    for i in range(lengh):
        if message.text in files[i]:
            file_name = message.text
        elif message.text == str(i + 1):
            file_name = files[i]
    if os.path.exists(folder_path + "/" + file_name):
        file = open("шаблон-файл.txt","a",encoding="utf-8")
        file.write(file_name + "-")
        await message.answer("Дайте название файлу")
        await state.set_state(createstate.waiting_for_choose_file_name)
    else:
        await message.answer(
            "Ошибка файла с " + message.text + " номером/названием не существует, проверьте правильность номера/названия введёного вами")

@dp.message(createstate.waiting_for_choose_file_name)
async def name_get(message: types.Message, state:FSMContext):
    dst = "Отчёты/Пример1.xlsx"
    shutil.copy("Отчёты/Пример.xlsx",dst)
    os.rename("Отчёты/Пример1.xlsx", "Отчёты/" + message.text)
    file = open("шаблон-файл.txt","a",encoding="utf-8")
    file.write(message.text + ".xlsx" + "\n")
    await message.answer("Файл успешно создан")
    await state.clear()