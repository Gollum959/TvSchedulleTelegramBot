import pendulum
import re
import os
import shutil
import telebot
import ScheduleDocument
import win32com.client as win32
from win32com.client import constants
#import Chekandcopy


today = pendulum.today()

list_dir = os.listdir(path="E:\\Work\\Python\\Project\\Test Directory")
work_dir = "E:\\Work\\Python\\Project\\Test Right Directory\\"
this_week_local_file_name = "thisweek"
next_week_local_file_name = "nextweek"

current_week_start = today.start_of('week')
current_week_end = today.end_of('week')
next_week_start = current_week_start.add(weeks=1)
next_week_end = current_week_end.add(weeks=1)


def regular_string(week_start, week_end):
    regular_string = r'(0' + str(week_start.day) + '|' + str(week_start.day) + ')\
[.]*\
(0' + str(week_start.month) + '|' + str(week_start.month) + ')*\
[.]*\
(' + str(week_start.year) + ')*\
[\s]*-[\s]*\
(0' + str(week_end.day) + '|' + str(week_end.day) + ')\
[.]*\
(0' + str(week_end.month) + '|' + str(week_end.month) + ')*\
[.]*\
(' + str(week_end.year) + ')*'
    return regular_string


def path(week_start, week_end):
    regular_strig_week = regular_string(week_start, week_end)
    for directory in list_dir:
        match_dir = re.findall(regular_strig_week, directory)
        if match_dir:
            return f'E:\Work\Python\Project\Test Directory\{directory}'
    return ''


def file_path(week_start, week_end):
    path_file = path(week_start, week_end)
    if path_file:
        list_file = os.listdir(path=path_file)
        for file in list_file:
            match_file = re.findall(r"(p|р)(т|t)(s)[\w \W]*(.doc|.docx|rtf)$", file, re.IGNORECASE)
            if match_file:
                return f'{path_file}\{file}'
    return ''


def chek_and_copy(path_to_file, local_file, text):
    extension = '.docx' if path_to_file.endswith('.docx') else ('.doc' if path_to_file.endswith('.doc') else '')
    source_file_info = os.stat(path_to_file)
    source_file_size = source_file_info.st_size
    local_file_path = work_dir+local_file+extension

    if os.path.isfile(local_file_path):
        local_file_info = os.stat(local_file_path)
        local_file_size = local_file_info.st_size
    else:
        local_file_size = 0

    if source_file_size == local_file_size:
        print(f'File {text} is already exist')
    else:
        shutil.copyfile(path_to_file, local_file_path, follow_symlinks=True)
        print(f'File {text} is copied')
        if extension == '.doc':
            save_as_docx(local_file_path, local_file)
            print(f'File {text} is converted')




def save_as_docx(path, file_name):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()

    # Rename path with .docx
    new_file_abs = work_dir+file_name+'.docx'

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)


def chek_file_exist(source_file_path, local_file_name, text):
    if os.path.isfile(source_file_path):
        chek_and_copy(source_file_path, local_file_name, text)
    else:
        print(f'Source file {text} isn`t found')


this_weak_source_file_path = file_path(current_week_start, current_week_end)
next_weak_source_file_path = file_path(next_week_start, next_week_end)

if this_weak_source_file_path:
    chek_file_exist(this_weak_source_file_path, this_week_local_file_name, 'for this week')
else:
    print(f'Source file for this week isn`t found')

if next_weak_source_file_path:
    chek_file_exist(next_weak_source_file_path, next_week_local_file_name, 'for next week')
else:
    print(f'Source file for next week isn`t found')

doc_this_week = ScheduleDocument.ScheduleDocument(work_dir+this_week_local_file_name+'.docx')
doc_this_week.init()
list_pts_and_hardware_this_week = doc_this_week.get_map()

doc_next_week = ScheduleDocument.ScheduleDocument(work_dir+next_week_local_file_name+'.docx')
doc_next_week.init()
#list_pts_and_hardware_next_week = doc_next_week.map_pts_and_hardware

token = '2133524927:AAGL7Quqshq3OaLmfmFwTpSqaqWj25g2p0g'
bot = telebot.TeleBot(token)

@bot.message_handler(commands=['start'])
def handle_text(message):
    user_markup = telebot.types.ReplyKeyboardMarkup(True, False)
    for name_hardware in list_pts_and_hardware_this_week:
        user_markup.row(name_hardware)
    bot.send_message(message.from_user.id, 'Добро пожаловать в бот по получению расписания ПТС!\n Выберите Пункт МЕНЮ', reply_markup=user_markup)

@bot.message_handler(func=lambda mess: mess.text, content_types=['text'])
def show_day_timetable(message):
    keyboard = telebot.types.InlineKeyboardMarkup()
    callback_data_menu_1 = {}
    callback_data_menu_1['id']=1
    callback_data_menu_1['message']=message.text
    menu_1 = telebot.types.InlineKeyboardButton(text='Расписание на сегодня', callback_data= "submenu_1_"+message.text)
    menu_2 = telebot.types.InlineKeyboardButton(text='Расписание на неделю', callback_data="submenu_2_"+message.text)
    menu_3 = telebot.types.InlineKeyboardButton(text='Расписание на Следующую неделю', callback_data="submenu_3_"+message.text)
    keyboard.add(menu_1, menu_2, menu_3)
    bot.send_message(message.chat.id, 'It works!', reply_markup=keyboard)

@bot.callback_query_handler(func=lambda call: call.data.startswith('submenu_'))
def show_day_timetable(call):
    splited_submenu = call.data.split('_')
    name_pts = splited_submenu[2]
    type_schedule = int(splited_submenu[1])
    if type_schedule == 1:
        result = doc_this_week.get_day_timetable_pts(name_pts)
        bot.send_message(call.from_user.id, result)
    elif type_schedule == 2:
        result = doc_this_week.get_week_timetable_pts(name_pts)
        bot.send_message(call.from_user.id, result)
    elif type_schedule == 3:
        result = doc_next_week.get_week_timetable_pts(name_pts)
        bot.send_message(call.from_user.id, result)

bot.polling(none_stop=True)

