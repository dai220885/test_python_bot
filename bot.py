import openpyxl
from token import API_TOKEN
from aiogram import Bot, Dispatcher, executor, types

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

path = 'testEXEL.xlsx'
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
contacts = list(sheet.values)

@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    await message.reply("Привет!\nЭто бот для работы с телефонной книгой")

@dp.message_handler(commands=['help'])
async def send_help(message: types.Message):
    await message.reply("Список доступных комманд:\n"
                        "\'/read\' - чтение всей телефонной книги\n"
                        "\'/find\' [критерий поиска] - поиск в телефонной книге по указанному критерию "
                                                        "(фамилия, имя, номер телефона и т.д.")


@dp.message_handler(content_types=['text'])
async def send_read(message: types.Message):

    if message.text.lower() == '/read':
        await message.answer(f'Контакты телефонной книги:')
        number = 1
        for contact in contacts:
            await message.answer(f'контакт номер {number}\nфамилия: {contact[0]}\nимя: {contact[1]}\nтелефон: {contact[2]}\n')
            number += 1

    elif message.text.lower()[:5] == '/find':
        if message.text[6:] == '' or message.text[6:] == ' ':
            await message.answer('уточните критерии поиска')
        else:
            await message.answer(f'результат поиска в телефонной книге: {message.text[6:]}')
            rezult = 0
            for contact in contacts:
                for item in contact:
                    if message.text[6:] in item:
                        await message.answer(f'фамилия: {contact[0]}\nимя: {contact[1]}\nтелефон: {contact[2]}\n')
                        rezult += 1
                        break
            if rezult == 0:
                await message.answer('совпадений не найдено')

    else:
        await message.answer('Не понимаю, что это значит. введите команду \'/help\' для получения списка доступных комманд')


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
