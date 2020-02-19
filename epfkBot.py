import json
import telebot
import xlrd
import requests

bot = telebot.TeleBot('1075655909:AAHQG3aYP3DHMXUaXdmTmbmMDcqzImRZTxE')

@bot.message_handler(commands=["start"])
def startMessage(message):
	keyboard = telebot.types.InlineKeyboardMarkup()

	button = telebot.types.InlineKeyboardButton(text="Розклад", callback_data=json.dumps({'course':'', 'group':'', 'day':0}))
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

		button = telebot.types.InlineKeyboardButton(text="Розклад", callback_data=json.dumps({'course':'', 'group':'', 'day':0}))
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
			button = telebot.types.InlineKeyboardButton(text=f"{courses[course]} курс", callback_data=json.dumps({'course':courses[course], 'group':0, 'day':0}))
			keyboard.add(button)
		button = telebot.types.InlineKeyboardButton(text="Назад", callback_data='MAINMENU')
		keyboard.add(button)
		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=f"Будь ласка, оберіть ваш курс",
			reply_markup=keyboard)

	elif jsonData['group'] == 0:
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
		counter = 0
		for col in sheet.row_values(row):
			if len(col) > 0 and col[0] == 'К':
				counter += 1
				newJsonData['group'] = counter
				button = telebot.types.InlineKeyboardButton(text=col, callback_data=json.dumps(newJsonData))
				keyboard.add(button)
		newJsonData['course'] = ''
		newJsonData['group'] = 0
		newJsonData['day'] = 0
		button = telebot.types.InlineKeyboardButton(text='Назад', callback_data=json.dumps(newJsonData))
		keyboard.add(button)

		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=f"Будь ласка, оберіть вашу спеціальність",
			reply_markup=keyboard)

	elif jsonData['day'] == 0:
		days = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', 'П\'ятниця']

		keyboard = telebot.types.InlineKeyboardMarkup()
		newJsonData = jsonData
		for day in range(0, 5):
			newJsonData['day'] = day + 1
			button = telebot.types.InlineKeyboardButton(text=days[day], callback_data=json.dumps(newJsonData))
			keyboard.add(button)

		newJsonData['group'] = newJsonData['day'] = 0
		button = telebot.types.InlineKeyboardButton(text='Назад', callback_data=json.dumps(newJsonData))
		keyboard.add(button)

		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=f"Будь ласка, оберіть день тижня",
			reply_markup=keyboard)

	elif jsonData['day'] != 0:
		days = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', 'П\'ятниця']
		messageDays = ['Понеділок', 'Вівторок', 'Середу', 'Четвер', 'П\'ятницю']
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

		counter = groupCounter = 0
		for col in sheet.row_values(row):
			counter += 1
			if len(col) > 0 and col[0] == 'К':
				groupCounter += 1
				if groupCounter == jsonData['group']: break

		counter -= 1
		GC = counter
		GR = row
		row = 1
		while sheet.row_values(row)[0].strip() != days[jsonData['day']-1].strip():
			row += 1

		lessonCounter = 0
		messageText = 'Розклад на ' + messageDays[jsonData["day"]-1].lower() + ' для групи ' + sheet.cell_value(GR, GC)[:sheet.cell_value(GR, GC).find('\n')] + '\n\n'
		while True:
			if len(str(sheet.cell_value(row, 2))) > 0 and sheet.cell_value(row, 2) > lessonCounter:
				lessonCounter = sheet.cell_value(row, 2)
			elif len(str(sheet.cell_value(row, 2))) == 0 or sheet.cell_value(row, 2) < lessonCounter:
				break

			messageText += f'{int(lessonCounter)}. '
			if len(sheet.cell_value(row, counter)) > 0:
				messageText += f'{sheet.cell_value(row, counter)} '
			else:
				messageText += '– '
			if len(sheet.cell_value(row+2, counter)) > 0 and sheet.cell_value(row+2, counter).find('.') == -1:
				messageText += f'| {sheet.cell_value(row+2, counter)} '
			if len(sheet.cell_value(row+2, counter)) == 0 and sheet.cell_value(row+1, counter).find('.') != -1:
				messageText += f'| –'

			messageText += '\n'

			row += 4

		keyboard = telebot.types.InlineKeyboardMarkup()
		newJsonData = jsonData
		newJsonData['day'] = 0
		button = telebot.types.InlineKeyboardButton(text='Назад', callback_data=json.dumps(newJsonData))
		keyboard.add(button)

		bot.edit_message_text(
			chat_id=call.message.chat.id, 
			message_id=call.message.message_id, 
			text=messageText,
			reply_markup=keyboard)

bot.polling()