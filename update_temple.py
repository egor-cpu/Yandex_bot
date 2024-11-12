from aiogram.types import ReplyKeyboardRemove

from main import dp, bot, y, greet_kb
import yadisk
from aiogram import types
import os
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram import types

def update_files_from_yandex_disk(local_dir, yandex_disk_dir):
    if not os.path.exists(local_dir):
        os.makedirs(local_dir)

    # Получение списка файлов на Яндекс.Диске
    for item in y.listdir(yandex_disk_dir):
        if item['type'] == 'file':
            file_name = item['name']
            remote_path = f"{yandex_disk_dir}/{file_name}"
            local_path = os.path.join(local_dir, file_name)

            remote_modified_time = item['modified'].timestamp()
            # Проверка необходимости загрузки/обновления файла
            if not os.path.exists(local_path) or remote_modified_time > os.path.getmtime(local_path):
                y.download(remote_path, local_path)
            else:
                print(f"Файл {file_name} уже актуален.")

class updateStates(StatesGroup):
    waiting_for_send_temple = State()
@dp.message(lambda message: message.text == "Обновить шаблоны")
async def update_button(message: types.Message, state:FSMContext):
    file = open("admins.txt", 'r')
    booli = 0
    for i in file:
        if i == str(message.chat.id) + "\n":
            booli=1
    if booli==0:
        await message.answer("У вас недостаточно доступа")
    else:
        keyboard = types.ReplyKeyboardMarkup(keyboard=[], resize_keyboard=True)
        keyboard.keyboard.append([types.KeyboardButton(text="Отмена")])
        remote_file_path = "/tg test/Шаблоны"
        local_file_path = "Шаблоны"
        update_files_from_yandex_disk(local_dir=local_file_path + "/",yandex_disk_dir=remote_file_path)
        await message.answer("Можете отправлять новый шаблон", reply_markup=ReplyKeyboardRemove())
        await message.answer("Добавлю кнопку для отмены", reply_markup=keyboard)
        await state.set_state(updateStates.waiting_for_send_temple)
@dp.message(updateStates.waiting_for_send_temple)
async def doc_get(message:types.Message,state:FSMContext):
    if message.text=="Отмена":
        await message.answer("Возвращаю вас к началу", reply_markup=greet_kb)
        await state.clear()
    else:
        document = message.document
        file_id = document.file_id
        file_info = await bot.get_file(file_id)
        file_path = file_info.file_path
        file_name = document.file_name
        file_path_to_save = os.path.join("Шаблоны", file_name)
        downloaded_file = await bot.download_file(file_path)
        with open(file_path_to_save, "wb") as f:
            f.write(downloaded_file.read())
        await message.answer("Загружаю на диск", reply_markup=ReplyKeyboardRemove())
        y.upload("Шаблоны/" + file_name, "/tg test/Шаблоны/" + file_name)
        await message.answer("Готово, возвращаю вас к началу",reply_markup=greet_kb)
        await state.clear()