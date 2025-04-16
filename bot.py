import gspread
from oauth2client.service_account import ServiceAccountCredentials
import telebot
from telebot.types import InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
import json
from gspread_formatting import *
from gspread_formatting import CellFormat, Color

# Вставьте сюда правильный токен для бота
bot = telebot.TeleBot('7732089016:AAFWBeDjKkK-jLibqf_pgE66LDibMPWMZYs')

# Учетные данные для доступа к Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
creds = ServiceAccountCredentials.from_json_keyfile_name('sotrydnufd.json', SCOPE)  # Путь к JSON-файлу с учетными данными
client = gspread.authorize(creds)

# Открытие документа Google Sheets по имени
sheet = client.open('Задача').sheet1  # Замените на имя своей Google Sheets таблицы

# Глобальный словарь для хранения ключей сотрудников
employee_sessions = {}

# Ожидаемые заголовки для всей таблицы
expected_headers = ['EmployeeKey', 'Task', 'Materials', 'Column6', 'Column7']

# Чтение данных из Google Sheets
def get_employee_tasks(employee_key):
    tasks = []
    records = sheet.get_all_records(expected_headers=expected_headers)  # Получаем все строки как словари с ожидаемыми заголовками
    for record in records:
        if record['EmployeeKey'] == employee_key and not record['Task'].endswith("Z"):  # Ищем задачи без буквы "Z" в конце
            tasks.append(record['Task'])  # Добавляем задачу в список
    return tasks

# Чтение данных о материалах из Google Sheets
def get_employee_materials(employee_key):
    materials = []
    records = sheet.get_all_records(expected_headers=expected_headers)  # Получаем все строки как словари с ожидаемыми заголовками
    for record in records:
        if record['EmployeeKey'] == employee_key:  # Проверяем, совпадает ли ключ сотрудника с данным
            materials.append(record['Materials'])  # Добавляем материал в список
    return materials

# Чтение данных из файла JSON
def load_employee_data():
    with open('employees.json', 'r', encoding='utf-8') as file:
        return json.load(file)

# Функция для добавления сотрудника в файл
def add_employee(employee_key, employee_name):
    employee_data = load_employee_data()  # Загружаем текущие данные
    employee_data[str(employee_key)] = employee_name  # Добавляем нового сотрудника
    
    # Записываем обновленные данные обратно в файл
    with open('employees.json', 'w', encoding='utf-8') as file:
        json.dump(employee_data, file, ensure_ascii=False, indent=4)

# Функция для создания кнопки "Попробовать еще"
def try_again_keyboard():
    keyboard = InlineKeyboardMarkup()
    button = InlineKeyboardButton(text="Попробовать еще", callback_data="try_again")
    keyboard.add(button)
    return keyboard

# Функция для создания кнопок "Задачи" и "Материалы"
def task_materials_keyboard():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=False)
    button1 = KeyboardButton("Задачи")
    button2 = KeyboardButton("Материалы")
    keyboard.add(button1, button2)
    return keyboard

# Функция для отображения задач как кнопок
def display_tasks(message, tasks):
    keyboard = InlineKeyboardMarkup()
    
    # Для каждой задачи создаем кнопку с уникальным callback_data
    for task in tasks:
        button = InlineKeyboardButton(text=task, callback_data=f"task_{task}")  # Используем саму задачу как уникальный идентификатор
        keyboard.add(button)

    bot.send_message(message.chat.id, "Выберите задачу:", reply_markup=keyboard)

# Функция для отображения материалов как кнопок
def display_materials(message, materials):
    keyboard = InlineKeyboardMarkup()
    
    # Для каждого материала создаем кнопку
    for i, material in enumerate(materials):
        button = InlineKeyboardButton(text=material, callback_data=f"material_{i}")
        keyboard.add(button)

    bot.send_message(message.chat.id, "Выберите материал:", reply_markup=keyboard)

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    # Создаем inline-кнопку "Начать"
    keyboard = InlineKeyboardMarkup()
    button = InlineKeyboardButton(text="Начать", callback_data="start_button")
    keyboard.add(button)
    
    # Отправляем сообщение с кнопкой
    bot.send_message(message.chat.id, "Я мастер по распределению задач от Klumneo дизайн. Помогаю не запутаться в задачах и в сотни excel файлах. По тех.вопросам или приобритение подписки @licvis Нажимай начать", reply_markup=keyboard)

# Обработчик нажатия на кнопку "Начать"
@bot.callback_query_handler(func=lambda call: call.data == "start_button")
def handle_start_button(call):
    # Отправляем картинку
    photo_url = 'https://artfedoseev.ru/wp-content/uploads/2025/03/fgdtrt-1024x673.png'  # Замените на ссылку на вашу картинку
    bot.send_photo(call.message.chat.id, photo_url)  # Отправка изображения
    
    # Просим ввести ключ
    bot.send_message(call.message.chat.id, "Перед началом работы необходимо ввести ключ сотрудника. Не переживай, кредит не оформим! Обратись к своему руководителю, если не понял о чем речь ")

    # Переходим к следующему шагу, ожидаем ввода ключа
    bot.register_next_step_handler(call.message, verify_employee_key)

# Обработчик нажатия на кнопку "Попробовать еще"
@bot.callback_query_handler(func=lambda call: call.data == "try_again")
def handle_try_again(call):
    # Запрашиваем ключ снова
    bot.send_message(call.message.chat.id, "Введите ваш ключ еще раз.")
    bot.register_next_step_handler(call.message, verify_employee_key)

# Функция для проверки ключа сотрудника
def verify_employee_key(message):
    employee_data = load_employee_data()  # Загружаем актуальные данные
    employee_key_str = message.text.strip()  # Получаем ключ как строку
    
    # Преобразуем ключ в целое число
    try:
        employee_key = int(employee_key_str)
    except ValueError:
        bot.send_message(message.chat.id, "Неверный формат ключа. Пожалуйста, введи числовой ключ.")
        return

    if str(employee_key) in employee_data:  # Сравниваем строковое значение ключа
        # Сохраняем ключ в сессии
        employee_sessions[message.chat.id] = employee_key

        # Отправляем приветствие сотруднику
        employee_name = employee_data[str(employee_key)]  # Строковый ключ для доступа в JSON
        bot.send_message(message.chat.id, f"Привет, {employee_name}! Поздравляю с успешной авторизацией. Ну че, поворкаем?")

        # Показываем кнопки "Задачи" и "Материалы"
        bot.send_message(message.chat.id, "Тебе доступно для выбора", reply_markup=task_materials_keyboard())
    else:
        # Если ключ неправильный, отправляем сообщение с кнопкой "Попробовать еще"
        bot.send_message(message.chat.id, "Похоже, ты не работаешь у нас. Доступ запрещен", reply_markup=try_again_keyboard())

# Обработчик выбора "Задачи" или "Материалы"
@bot.message_handler(func=lambda message: message.text in ["Задачи", "Материалы"])
def handle_task_materials_selection(message):
    employee_key = employee_sessions.get(message.chat.id)

    if employee_key:
        if message.text == "Задачи":
            tasks = get_employee_tasks(employee_key)
            if tasks:
                display_tasks(message, tasks)
            else:
                bot.send_message(message.chat.id, "Кажется, у вас нет активных задач. Обратись к своему руководителю")
        elif message.text == "Материалы":
            materials = get_employee_materials(employee_key)
            if materials:
                display_materials(message, materials)
            else:
                bot.send_message(message.chat.id, "Кажется у тебя лайт работа и тебе не нужны материалы")
    else:
        bot.send_message(message.chat.id, "Ошибка: не удалось найти твою сессию. Напиши в чат /start и пройди авторизацию")

# Обработчик выбора задачи
@bot.callback_query_handler(func=lambda call: call.data.startswith('task_'))
def handle_task_selection(call):
    task = call.data.split('_')[1]  # Получаем уникальный идентификатор задачи
    employee_key = employee_sessions.get(call.message.chat.id)

    if employee_key:
        # Найдем строку с нужной задачей в Google Sheets
        records = sheet.get_all_records(expected_headers=expected_headers)  # Получаем все записи из таблицы
        
        for i, record in enumerate(records):
            if record['Task'] == task and record['EmployeeKey'] == employee_key:
                column_6 = record.get('Column6', 'Нет данных')  # Имя столбца в таблице должно быть корректным
                column_7 = record.get('Column7', 'Нет данных')  # Имя столбца в таблице должно быть корректным
                
                # Отправляем информацию о задаче
                bot.send_message(call.message.chat.id, f"Задача: {task}\nО задаче: {column_6}\nМатериалы: {column_7}")
                
                # Создаем клавиатуру с кнопкой "Завершить"
                keyboard = InlineKeyboardMarkup()
                finish_button = InlineKeyboardButton(text="Завершить", callback_data=f"finish_{task}")
                keyboard.add(finish_button)

                bot.send_message(call.message.chat.id, "Не забудь нажать кнопку чтобы завершить проект и приступить к новому", reply_markup=keyboard)
                break
        else:
            bot.send_message(call.message.chat.id, "Задача не найдена.")

    else:
        bot.send_message(call.message.chat.id, "Ошибка: не удалось найти твою сессию. Напиши в чат /start и пройди авторизацию")

# Обработчик завершения задачи
@bot.callback_query_handler(func=lambda call: call.data.startswith('finish_'))
def handle_finish_task(call):
    task = call.data.split('_')[1]  # Получаем уникальный идентификатор задачи
    employee_key = employee_sessions.get(call.message.chat.id)

    if employee_key:
        # Обновление записи в Google Sheets, чтобы пометить задачу как завершенную
        records = sheet.get_all_records(expected_headers=expected_headers)
        
        # Ищем задачу в таблице
        for i, record in enumerate(records):
            if record['Task'] == task and record['EmployeeKey'] == employee_key:
                # Обновляем колонку с задачей
                sheet.update_cell(i + 2, 2, f"{task}Z")  # Здесь '2' - это номер строки, а '2' - это столбец Task

                # Помечаем ячейку зеленым цветом
                format_cell_range(sheet, f'B{i+2}', CellFormat(backgroundColor=Color(0, 1, 0)))  # B - это столбец с задачей
                
                bot.send_message(call.message.chat.id, f"Задача '{task}' была завершена успешно!")
                break
        else:
            bot.send_message(call.message.chat.id, "Задача не найдена.")
    else:
        bot.send_message(call.message.chat.id, "Ошибка: не удалось найти твою сессию. Напиши в чат /start и пройди авторизацию")

# Запуск бота
bot.polling(none_stop=True, interval=0)
