import openpyxl
from main import dp, bot, y
import yadisk
from aiogram import Bot, Dispatcher, types
from dotenv import load_dotenv
from aiogram import types
from aiogram.filters import Command
import os
from aiogram.types import ReplyKeyboardRemove, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
class editStates(StatesGroup):
    waiting_for_chose_document = State()
    waiting_for_worktype = State()
    waiting_for_chose_row = State()
    waiting_for_element_row = State()
    waiting_for_edit_row = State()
    waiting_for_chose_date = State()
    waiting_for_chose_category = State()
    waiting_for_chose_expense_item = State()
    waiting_for_chose_name = State()
    waiting_for_type = State()
    waiting_for_chose_count = State()
    waiting_for_chose_price = State()
    waiting_for_document_numbers = State()
    waiting_for_document_types = State()
    waiting_for_document = State()
    waiting_for_comments = State()

    waiting_for_accept_row = State()

filenamesave = ""
allrow = ""
numberline=0
@dp.message(lambda message: message.text == "Изменить Отчёт")
async def exel_edit(message: types.Message, state: FSMContext):
    file = open("admins.txt", 'r')
    booli = 0
    for i in file:
        if i == str(message.chat.id):
            booli=1
    if booli==0:
        await message.answer("У вас недостаточно доступа")
        await state.clear()
    else:
        files = os.listdir("Отчёты")
        global lengh
        a = len(files)
        lengh = a
        print(lengh)
        await message.answer("Выберите отчёт\n")
        for i in range(a):
            await message.answer(str(i + 1) + ". " + str(files[i]))
        await state.set_state(editStates.waiting_for_chose_document)

@dp.message(editStates.waiting_for_chose_document)
async def chose_document_get(message: types.Message, state: FSMContext):
    folder_path = 'Отчёты'
    files = os.listdir("Отчёты")
    global filenamesave
    file_name = message.text
    for i in range(lengh):
        if message.text in files[i]:
            file_name = message.text
        elif message.text == str(i+1):
            file_name = files[i]
    if os.path.exists(folder_path + "/" + file_name):
        filenamesave = file_name
        await message.answer("Выберите режим работы: \n 1. Изменить строку в отчёте \n 2. Добавить строку ")
        await state.set_state(editStates.waiting_for_worktype)
    else:
        await message.answer("Ошибка файла с " + message.text + " номером/названием не существует, проверьте правильность номера/названия введёного вами")

@dp.message(editStates.waiting_for_worktype)
async def worktype_get(message: types.Message, state: FSMContext):
    if message.text in "Изменить строку в отчёте" or message.text == str(1):
        await message.answer("Напишите номер строки")
        await state.set_state(editStates.waiting_for_chose_row)
    elif message.text in "Добавить строку" or message.text == str(2):
        await message.answer("Введите дату покупки")
        await state.set_state(editStates.waiting_for_chose_date)
    else:
        await message.answer("Вы неправильно ввели команду")
        await state.set_state(editStates.waiting_for_worktype)


@dp.message(editStates.waiting_for_chose_row)
async def chose_row_get(message:types.Message, state:FSMContext):
    global numberline
    numberline = int(message.text)
    await message.answer("Напишите название столбца (A или I)")
    await state.set_state(editStates.waiting_for_element_row)

@dp.message(editStates.waiting_for_element_row)
async def chose_element_get(message:types.Message, state:FSMContext):
    global allrow
    allrow=message.text
    await message.answer("Напишите текст")
    await state.set_state(editStates.waiting_for_edit_row)
@dp.message(editStates.waiting_for_edit_row)
async def edit_row_get(message:types.Message, state:FSMContext):
    wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
    sheet = wb.active
    A = sheet[allrow + str(numberline)]
    A.value = message.text
    wb.save("Отчёты/" + filenamesave)
    await message.answer("Строка изменена. \nНапишите Да для сохранения изменений")
    await state.set_state(editStates.waiting_for_accept_row)
@dp.message(editStates.waiting_for_chose_date)
async def date_gate(message:types.Message, state:FSMContext):
    if message.text == "Выход":
        await message.answer("Вы вышли")
        await state.clear()
    else:
        global filenamesave
        global numberline
        global allrow
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        i = 1
        A = sheet['A' + str(i)]
        while A.value != None:
            print(sheet['A' + str(i)])
            A = sheet['A' + str(i)]
            i+=1
        numberline = i
        A.value = message.text
        allrow="Дата:" + message.text + "\n"
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Выберите категорию:")
        file = open("typecategory.txt", "r")
        i = 1
        text = ""
        for string in file:
            text = text + (str(i)+". " + string)
            i+=1
        await message.answer(text)
        file.close()
        await state.set_state(editStates.waiting_for_chose_category)

@dp.message(editStates.waiting_for_chose_category)
async def category_get(message:types.Message, state: FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_date)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        file = open("typecategory.txt", "r")
        i = 1
        text = "none"
        for string in file:
            if (message.text == str(i)) or message.text in string:
                text=string
            i += 1
        file.close()
        if text != "none":
            wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
            sheet = wb.active
            B = sheet['B' + str(numberline)]
            B.value = text
            global allrow
            allrow = allrow + "Категория:" + text + "\n"
            wb.save("Отчёты/" + filenamesave)
            await message.answer("Статью расходов:")
            file = open("expence_item.txt", "r")
            i = 1
            text = ""
            for string in file:
                text = text + (str(i) + ". " + string)
                i += 1
            await message.answer(text)
            file.close()
            await state.set_state(editStates.waiting_for_chose_expense_item)
        else:
            await message.answer("Вы неправильно ввели категорию, попробуйте ещё раз")
            await state.set_state(editStates.waiting_for_chose_category)

@dp.message(editStates.waiting_for_chose_expense_item)
async def expencse_get(message:types.Message, state: FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_category)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        file = open("expence_item.txt", "r")
        i = 1
        text = "none"
        for string in file:
            if (message.text == str(i)) or message.text in string:
                text = string
            i += 1
        file.close()
        if text != "none":
            wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
            sheet = wb.active
            C = sheet['C' + str(numberline)]
            global allrow
            allrow = allrow + "Статья расходов:" + text + "\n"
            C.value = text
            wb.save("Отчёты/" + filenamesave)
            await message.answer("Напишите название")
            await state.set_state(editStates.waiting_for_chose_name)
        else:
            await message.answer("Вы неправильно ввели Статью расходов, попробуйте ещё раз")
            await state.set_state(editStates.waiting_for_chose_expense_item)

@dp.message(editStates.waiting_for_chose_name)
async def name_get(message:types.Message, state:FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_expense_item)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        D = sheet['D' + str(numberline)]
        global allrow
        allrow = allrow + "Тип:" + message.text + "\n"
        D.value = message.text
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Какой тип: \n 1. Грантовый \n 2. Локальный")
        await state.set_state(editStates.waiting_for_type)

@dp.message(editStates.waiting_for_type)
async def type_get(message:types.Message,state:FSMContext):
    global allrow
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_name)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        if message.text == "1" or message.text == "Грантовый":
            wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
            sheet = wb.active
            E = sheet['E' + str(numberline)]
            allrow = allrow + "Тип:" + "Грантовый" + "\n"
            E.value = message.text
            wb.save("Отчёты/" + filenamesave)
            await message.answer("Введите колличество")
            await state.set_state(editStates.waiting_for_chose_count)
        elif message.text == "2" or message.text == "Локальный":
            wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
            sheet = wb.active
            E = sheet['E' + str(numberline)]
            allrow = allrow + "Тип:" + "Локальный" + "\n"
            E.value = message.text
            wb.save("Отчёты/" + filenamesave)
            await message.answer("Введите колличество")
            await state.set_state(editStates.waiting_for_chose_count)
        else:
            await message.answer("Вы неправильно ввели тип, попробуйте ещё раз")
            await state.set_state(editStates.waiting_for_type)

@dp.message(editStates.waiting_for_chose_count)
async def count_get(message:types.Message, state: FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_type)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        F = sheet['F' + str(numberline)]
        global allrow
        allrow = allrow + "Количество:" + message.text + "\n"
        F.value = message.text
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Напишите сумму")
        await state.set_state(editStates.waiting_for_chose_price)

@dp.message(editStates.waiting_for_chose_price)
async def price_get(message:types.Message, state:FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_count)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        G = sheet['G' + str(numberline)]
        global allrow
        allrow = allrow + "Сумма:" + message.text + "\n"
        G.value = message.text
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Напишите колличество документов")
        await state.set_state(editStates.waiting_for_document_numbers)

document_number = 0
document_count = 0
@dp.message(editStates.waiting_for_document_numbers)
async def number_get(message:types.Message, state:FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_chose_price)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        global document_number
        document_number=int(message.text)
        await message.answer("Напишите названия документов")
        await state.set_state(editStates.waiting_for_document_types)
@dp.message(editStates.waiting_for_document_types)
async def types_get(message:types.Message, state:FSMContext):
    if message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_document_numbers)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        H = sheet['H' + str(numberline)]
        global allrow
        allrow = allrow + "Документы:" + message.text + "\n"
        H.value = message.text
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Можете начать отправлять документы, только по одному, а то я могу не успеть:(")
        await state.set_state(editStates.waiting_for_document)
@dp.message(editStates.waiting_for_document)
async def document_get(message:types.Message, state:FSMContext):
    global document_count
    if document_number==document_count:
        await message.answer("Все файлы загружены, есть какие-то комментарии?\n Если нет поставьте -")
        await state.set_state(editStates.waiting_for_comments)
    elif message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_document_types)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    else:
        document_count+=1
        document = message.document
        file_id = document.file_id
        file_info = await bot.get_file(file_id)
        file_path = file_info.file_path
        download_dile = await bot.download_file(file_path)
        file_name = document.file_name
        file_path_to_save = os.path.join("Документы", file_name)
        downloaded_file = await bot.download_file(file_path)
        with open(file_path_to_save, "wb") as f:
            f.write(downloaded_file.read())
        y.upload("Документы/" + file_name, "/tg test/Документы/" +file_name, overwrite=True)
        public_url = y.get_download_link("/tg test/Документы/" + file_name)
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        I = sheet['I' + str(numberline)]
        global allrow
        allrow = allrow + "Документы:" + public_url + "\n"
        if I.value != None:
            I.value = str(I.value) + public_url + "\n"
        else:
            I.value = public_url + "\n"
        wb.save("Отчёты/" + filenamesave)
        if document_number== document_count:
            await message.answer("Все файлы загружены, есть какие-то комментарии?\n Если нет поставьте -")
            await state.set_state(editStates.waiting_for_comments)
        else:
            await message.answer("Принял, ожидайте")
            await message.answer("Присылайте следующий документ")
            await state.set_state(editStates.waiting_for_document)
@dp.message(editStates.waiting_for_comments)
async def comments_get(message:types.Message, state: FSMContext):
    global allrow
    if message.text == "-":
        await message.answer(allrow)
        await message.answer("Всё верно? \nНапишите Да, если всё ок, либо вернитесь и справьте нужную строку, либо напишите Выход, чтобы удалить последние изменения в файле")
        await state.set_state(editStates.waiting_for_accept_row)
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    elif message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_document)
    else:
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        I = sheet['I' + str(numberline)]
        allrow = allrow + "Комментарии:" + message.text
        I.value = str(I.value) + message.text
        wb.save("Отчёты/" + filenamesave)
        await message.answer(allrow)
        await message.answer("Всё верно? \nНапишите Да, если всё ок, либо вернитесь и справьте нужную строку, либо напишите Выход, чтобы удалить последние изменения в файле")
        await state.set_state(editStates.waiting_for_accept_row)
@dp.message(editStates.waiting_for_accept_row)
async def accept_get(message:types.Message, state: FSMContext):
    if message.text == "Да":
        y.upload("Отчёты/" + filenamesave, "/tg test/Отчёты/" + filenamesave, overwrite=True)
        await message.answer("Супер, файлы загружены на диск")
        await state.clear()
    elif message.text == "Выход":
        wb = openpyxl.load_workbook("Отчёты/" + filenamesave)
        sheet = wb.active
        sheet.delete_rows(numberline)
        wb.save("Отчёты/" + filenamesave)
        await message.answer("Всё удалено, кроме документов на яндекс диске, пожалуйста не забудьте их удалить")
        await state.clear()
    elif message.text=="/back":
        await message.answer("Вы вернулись на один шаг назад")
        await state.set_state(editStates.waiting_for_comments)

