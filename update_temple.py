from main import dp, bot, y
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
                print(f"Загружаем {file_name}...")
                y.download(remote_path, local_path)
                print(f"Файл {file_name} обновлен.")
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
        remote_file_path = "/tg_test/Шаблоны/"
        local_file_path = "Шаблоны"
        if y.exists(remote_file_path):
            update_files_from_yandex_disk(local_file_path,remote_file_path)
            await message.answer("Можете отправлять новый шаблон")
            await state.set_state(updateStates.waiting_for_send_temple)
        else:
            await message.answer(f"Файл {remote_file_path} не найден на Яндекс.Диске")

@dp.message(updateStates.waiting_for_send_temple)
async def doc_get(message:types.Message,state:FSMContext):
    document = message.document
    file_id = document.file_id
    file_info = await bot.get_file(file_id)
    file_path = file_info.file_path
    download_dile = await bot.download_file(file_path)
    file_name = document.file_name
    file_path_to_save = os.path.join("Шаблоны", file_name)
    y.upload("Шаблоны/" + file_name, "/tg test/Шаблоны/" + file_name)