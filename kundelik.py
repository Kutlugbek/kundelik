from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import logging

from aiogram import Bot, Dispatcher, executor, types

import openpyxl
from openpyxl import load_workbook
wb = load_workbook('table_kun.xlsx')
sheet = wb.active

API_TOKEN = '5071987654:AAGRTkdiZCp37hc-ueCq1q4FvEUcLA5K2JQ'


logging.basicConfig(level=logging.INFO)

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)


@dp.message_handler()
async def send_message(msg: types.Message, message=None):
	await bot.send_message(934436280, message)


options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
driver.get("https://login.kundelik.kz/login")

for i in range(1, 47):
	Aid_v = "A" + str(i)
	Bid_v = "B" + str(i)

	login_value = sheet[str(Aid_v)].value
	login = driver.find_element_by_name('login') 
	login.send_keys(login_value)

	password_value = sheet[str(Bid_v)].value
	password = driver.find_element_by_name('password')
	password.send_keys(password_value)

	try:
		submit = driver.find_element(By.XPATH, "//input[@type='submit']")
		submit.click()

		try:
			marks = driver.find_element(By.XPATH, "//a[@href='https://schools.kundelik.kz/marks.aspx']")
			marks.click()
		except Exception as e:
			marks = driver.find_element(By.XPATH, "//a[@href='https://schools.kundelik.kz/children/marks.aspx']")
			marks.click()


		logout = driver.find_element(By.NAME, 'logout')
		logout.click()

		driver.get("https://login.kundelik.kz/login")
		print(login_value)
	except Exception as e:
		message = login_value
		send_message(message)

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
