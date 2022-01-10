import telebot
import ScheduleDocument

this_week_local_file_name = "E:\\Work\\Python\\Project\\Test Right Directory\\thisweek.docx"
next_week_local_file_name = "E:\\Work\\Python\\Project\\Test Right Directory\\nextweek.docx"

doc_this_week = ScheduleDocument.ScheduleDocument(this_week_local_file_name)
doc_this_week.init()
list_pts_and_hardware_this_week = doc_this_week.get_map()

doc_next_week = ScheduleDocument.ScheduleDocument(next_week_local_file_name)
doc_next_week.init()

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
    callback_data_menu_1['id'] = 1
    callback_data_menu_1['message'] = message.text
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

