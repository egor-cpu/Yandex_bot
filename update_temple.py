from main import dp, bot, y
import yadisk
from aiogram import types

@dp.message(lambda message: message.text == "Обновить шаблоны")
async def update_button(message: types.Message):
    file = open("admins.txt", 'r')
    booli = 0
    for i in file:
        if i == str(message.chat.id):
            booli=1
    if booli==0:
        await message.answer("У вас недостаточно доступа")
    else:
        remote_file_path = "/tg_test/Шаблоны/"
        local_file_path = "Шаблоны"
        if y.exists(remote_file_path):
            with y.download(remote_file_path) as remote_file:
                with open(local_file_path, "wb") as local_file:
                    # Копируем данные из удаленного файла в локальный
                    local_file.write(remote_file.read())
        else:
            await message.answer(f"Файл {remote_file_path} не найден на Яндекс.Диске")