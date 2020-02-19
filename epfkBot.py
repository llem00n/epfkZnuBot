import json
import telebot
import xlrd
import requests

bot = telebot.TeleBot('1075655909:AAHQG3aYP3DHMXUaXdmTmbmMDcqzImRZTxE')

@bot.message_handler(commands=["start"])
def startMessage(message):
	keyboard = telebot.types.InlineKeyboardMarkup()

	button = telebot.types.InlineKeyboardButton(text="Розклад", callback_data=json.dumps({'course':'', 'group':'', 'day':''}))
	keyboard.add(button)
	button = telebot.types.InlineKeyboardButton(text='Сайт ЕПФК ЗНУ', url='https://epkznu.com/')
	keyboard.add(button)

	bot.send_message(
		message.chat.id, 
		"Це бот, за допомогою якого ви зможете дізнатися розклад занять в коледжі ЕПФК ЗНУ.", 
		reply_markup=keyboard)

@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
	if call.data == 'MAINMENU':
		keyboard = telebot.types.InlineKeyboardMarkup()

		button = telebot.types.InlineKeyboardButton(text="Розклад", callback_data=json.dumps({'course':'', 'group':'', 'day':''}))
		keyboard.add(button)
		button = telebot.types.InlineKeyboardButton(text='Сайт ЕПФК ЗНУ', url='https://epkznu.com/')
		keyboard.add(button)

		bot.edit_message_text(
			chat_id=call.message.chat.id,
			message_id=call.message.message_id,
			text="Це бот, за допомогою якого ви зможете дізнатися розклад занять в коледжі ЕПФК ЗНУ.", 
			reply_markup=keyboard)
		return None
	jsonData = json.loads(call.data)

	if jsonData['course'] == '':
		courses = ['I', 'II', 'III', 'IV']
		keyboard = keyboard = telebot.types.InlineKeyboardMarkup()
		for course in range(0, 4):
			button = telebot.types.InlineKeyboardButton(text=f"{courses[course]} курс", callback_data=json.dumps({'course':courses[course], 'group':'', 'day':''}))
			keyboard.add(button)
		button = telebot.types.InlineKeyboardButton(text="Назад", callback_data='MAINMENU')
		keyboard.add(button)
		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=f"Будь ласка, оберіть ваш курс",
			reply_markup=keyboard)

	elif jsonData['group'] == '':
		url = 'http://epkznu.com/розклад-занять/'
		page = requests.get(url).text

		startAddress = page[:page.find(f">Розклад {jsonData['course']} курсу<")].rfind("http")
		endAddress = page[startAddress:].find("\">") + startAddress

		file = open('sheet.xls', 'wb')
		file.write(requests.get(page[startAddress:endAddress]).content)
		file.close()
		wb = xlrd.open_workbook('sheet.xls')
		sheet = wb.sheet_by_index(0)

		row = 1
		while sheet.row_values(row)[0].find('День') == -1:
			row += 1

		keyboard = telebot.types.InlineKeyboardMarkup()
		newJsonData = jsonData
		for col in sheet.row_values(row):
			if len(col) > 0 and col[0] == 'К':
				newJsonData['group'] = col[:col.find('\n')]
				print(newJsonData)
				button = telebot.types.InlineKeyboardButton(text=col, callback_data=json.dumps(newJsonData))
				keyboard.add(button)
		newJsonData['course'] = newJsonData['group'] = ''
		button = telebot.types.InlineKeyboardButton(text='Назад', callback_data=json.dumps(newJsonData))
		keyboard.add(button)

		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=f"Будь ласка, оберіть вашу спеціальність",
			reply_markup=keyboard)



bot.polling()
