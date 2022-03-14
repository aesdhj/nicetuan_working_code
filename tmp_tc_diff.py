import pandas as pd
import os
pd.options.display.max_columns = None
pd.options.display.max_rows = None
from datetime import datetime, timedelta, date
import sys
import warnings
warnings.filterwarnings('ignore')
from appium import webdriver as appium_webdriver
from selenium import webdriver as selenium_webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import time
import shutil


tc_short_action, arrival_time_action, arrival_time_action_new = 1, 0, 1
tc_short_check_start, tc_short_check_end = 0, 0
arrival_time_start, arrival_time_end = 0, 0


def get_verification_code():
	server = 'http://localhost:4723/wd/hub'
	desired_caps = {
		"platformName": "Android",
		"deviceName": "KB2000",
		"appPackage": "com.oneplus.mms",
		"appActivity": "com.android.mms.ui.ConversationList",
		"noReset": True,
		"automationName": "UIAutomator2"
	}
	driver = appium_webdriver.Remote(server, desired_caps)
	wait = WebDriverWait(driver, 30)
	verification_texts = wait.until(EC.presence_of_all_elements_located((By.ID, 'com.oneplus.mms:id/conversation_snippet')))
	for item in verification_texts:
		verification_text = item.get_attribute('text')
		verification_code = re.search('\d{6}', verification_text).group(0)
		if '【十荟团】' in verification_text:
			return verification_code
			break


def find_element_by_clickable(item):
	element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, item)))
	return element


def find_element_by_presence(item):
	element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, item)))
	return element


def find_elements_by_presence(item):
	elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, item)))
	return elements


def date_select(date_start, date_end):
	date_panel_left_head_text = find_element_by_presence(
		'div.el-date-range-picker__content.is-left div.el-date-range-picker__header').text
	date_panel_left_head_text_month = int(re.search('(.*?) 年 (.*?) 月', date_panel_left_head_text).group(2))
	print(date_panel_left_head_text_month)
	date_panel_left_forward = find_element_by_presence(
		'div.el-date-range-picker__content.is-left div.el-date-range-picker__header button.el-icon-arrow-left')
	date_panel_left_back = find_element_by_presence(
		'div.el-date-range-picker__content.is-right div.el-date-range-picker__header button.el-icon-arrow-right')

	if date_start.month < date_panel_left_head_text_month:
		date_panel_left_forward.click()
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.available')
		for item in date_panel_left_table:
			if int(item.text) == date_start.day:
				item.click()
				break
		date_panel_left_back.click()
	elif date_start.month == date_panel_left_head_text_month:
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.available')
		for item in date_panel_left_table:
			if int(item.text) == date_start.day:
				item.click()
				break
	else:
		print('date_start_error')
		sys.exit(0)

	if date_end.month == date_panel_left_head_text_month:
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.available')
		for item in date_panel_left_table:
			if int(item.text) == date_end.day:
				item.click()
				break
	elif date_end.month > date_panel_left_head_text_month:
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.next-month')
		for item in date_panel_left_table:
			if int(item.text) == date_end.day:
				item.click()
				break
	elif date_end.month < date_panel_left_head_text_month:
		date_panel_left_forward.click()
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.available')
		for item in date_panel_left_table:
			if int(item.text) == date_end.day:
				item.click()
				break
		# date_panel_left_back.click()
	else:
		print('date_end_error')
		sys.exit(0)

if tc_short_action == 1:
	browser = selenium_webdriver.Chrome()
	browser.get('https://wms.nicetuan.net/login')
	wait = WebDriverWait(browser, 20)
	user = find_element_by_presence('input.input_all.name').send_keys('')
	code_button = find_element_by_clickable('button#code').click()
	verification_code = get_verification_code()
	print(verification_code)
	password = find_element_by_presence('input.input_all_code.password').send_keys(verification_code)
	warehouse_select = find_element_by_clickable('a.textbox-icon.combo-arrow').click()
	warehouse = find_element_by_clickable('div#_easyui_combobox_i1_4').click()
	login = find_element_by_clickable('a#login.button.hide').click()
	# 20210604
	time.sleep(1)
	panel_headers = find_elements_by_presence('div.panel-title.panel-with-icon')
	for item in panel_headers:
		if '提示' in item.text:
			close = find_elements_by_presence('a.panel-tool-close')[1].click()
			break
	time.sleep(1)
	stock_mangement = find_element_by_clickable('div.accordion-header-selected').click()
	warehouse_mangement = find_element_by_clickable('div.panel-last').click()

	for items in [
		# ('hanzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(2)'),
		# ('wenzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(3)'),
		('jiangsu', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(1)'),
		# ('xuzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(5)'),
		# ('hefei', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(6)'),
	]:
		lvyuezhongxin = find_element_by_clickable('div#_easyui_tree_19 span.tree-title').click()
		browser.switch_to.window(browser.window_handles[1])
		tc_mangement = find_element_by_clickable('ul[role="menubar"] span:nth-child(8) > li').click()
		tc_short_check = find_element_by_clickable(
			'ul[role="menubar"] span:nth-child(8) > li li.el-menu-item:nth-child(4)').click()
		station_button = find_element_by_clickable('div.el-form-item--small:nth-child(1) div.el-form-item__content').click()
		time.sleep(1)
		station_select = find_element_by_clickable(items[1])
		print(station_select.text)
		station_select.click()

		date_button = find_element_by_presence('i.el-range__icon').click()
		time.sleep(1)
		date_start = date.today() - timedelta(days=tc_short_check_start)
		date_end = date.today() - timedelta(days=tc_short_check_end)
		print(date_start, date_end)
		date_select(date_start, date_end)

		search_button = find_element_by_clickable('button.el-icon-search').click()
		while True:
			data_table = find_element_by_presence('div.el-table__body-wrapper')
			if data_table.text != '暂无数据':
				break
			time.sleep(1)
		downnload_button = find_element_by_clickable('button.el-icon-download').click()
		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'C:\Users\aesdhj\Desktop'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('服务站差异' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, '服务站差异{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(1)
		time.sleep(1)
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
	browser.quit()
else:
	pass

time.sleep(1)

if arrival_time_action == 1:
	browser = selenium_webdriver.Chrome()
	browser.get('https://wms.nicetuan.net/login')
	wait = WebDriverWait(browser, 20)
	user = find_element_by_presence('input.input_all.name').send_keys('')
	code_button = find_element_by_clickable('button#code').click()
	verification_code = get_verification_code()
	print(verification_code)
	password = find_element_by_presence('input.input_all_code.password').send_keys(verification_code)
	warehouse_select = find_element_by_clickable('a.textbox-icon.combo-arrow').click()
	warehouse = find_element_by_clickable('div#_easyui_combobox_i1_4').click()
	login = find_element_by_clickable('a#login.button.hide').click()
	# 20210604
	time.sleep(1)
	panel_headers = find_elements_by_presence('div.panel-title.panel-with-icon')
	for item in panel_headers:
		if '提示' in item.text:
			close = find_elements_by_presence('a.panel-tool-close')[1].click()
			break
	time.sleep(1)
	stock_mangement = find_element_by_clickable('div.accordion-header-selected').click()
	warehouse_mangement = find_element_by_clickable('div.panel-last').click()

	for items in [
		('hanzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(2)'),
		('wenzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(3)'),
		('jiangsu', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(4)'),
		('xuzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(5)'),
		('hefei', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(6)'),
	]:
		lvyuezhongxin = find_element_by_clickable('div#_easyui_tree_19 span.tree-title').click()
		browser.switch_to.window(browser.window_handles[1])
		tc_mangement = find_element_by_clickable('ul[role="menubar"] span:nth-child(8) > li').click()
		tc_arrival_time = find_element_by_clickable(
			'ul[role="menubar"] span:nth-child(8) > li li.el-menu-item:nth-child(6)').click()
		station_button = find_element_by_clickable('div.el-input--suffix:nth-child(1)').click()
		station_select = find_element_by_clickable(items[1])
		print(station_select.text)
		station_select.click()

		date_button = find_elements_by_presence('i.el-range__icon')[0].click()
		time.sleep(1)
		date_start = date.today() - timedelta(days=arrival_time_start)
		date_end = date.today() - timedelta(days=arrival_time_end)
		print(date_start, date_end)
		date_select(date_start, date_end)

		search_button = find_element_by_clickable('button.el-icon-search').click()
		while True:
			data_table = find_element_by_presence('div.el-table__body-wrapper')
			if data_table != '暂无数据':
				break
			time.sleep(1)
		downnload_button = find_element_by_clickable('button.el-icon-download').click()
		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'C:\Users\aesdhj\Desktop'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('服务站支线' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, '服务站支线{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(1)
		time.sleep(1)
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
	browser.quit()
else:
	pass


if arrival_time_action_new == 1:
	browser = selenium_webdriver.Chrome()
	browser.get('https://wms.nicetuan.net/login')
	wait = WebDriverWait(browser, 20)
	user = find_element_by_presence('input.input_all.name').send_keys('13770995264')
	code_button = find_element_by_clickable('button#code').click()
	verification_code = get_verification_code()
	print(verification_code)
	password = find_element_by_presence('input.input_all_code.password').send_keys(verification_code)
	time.sleep(1)
	warehouse_select = find_element_by_clickable('a.textbox-icon.combo-arrow').click()
	warehouse = find_element_by_clickable('div#_easyui_combobox_i1_4').click()
	login = find_element_by_clickable('a#login.button.hide').click()
	# 20210604
	time.sleep(1)
	panel_headers = find_elements_by_presence('div.panel.window.panel-htop')
	for item in panel_headers:
		if '提示' in item.text:
			close = item.find_element_by_css_selector('a.panel-tool-close').click()
	time.sleep(1)
	stock_mangement = find_element_by_clickable('div.accordion-header-selected').click()
	warehouse_mangement = find_element_by_clickable('div.panel-last').click()

	for items in [
		# ('hanzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(2)'),
		# ('wenzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(3)'),
		('jiangsu', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(4)'),
		# ('xuzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(5)'),
		# ('hefei', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(6)'),
	]:
		tms_center = find_element_by_clickable('div#_easyui_tree_20 span.tree-title').click()
		browser.switch_to.window(browser.window_handles[1])
		for item in find_elements_by_presence('ul.sideBox li'):
			if '支线管理' in item.text:
				item.click()
				time.sleep(1)
				dispatch_task_report = item.find_element_by_css_selector('ul[role="menu"] li:nth-child(1)').click()
				time.sleep(1)
				break
		warehouse_button = find_element_by_clickable('div.el-row--flex div.el-form-item--small:nth-child(1) span.el-input__suffix').click()
		warehouse_select = find_element_by_clickable(items[1])
		print(warehouse_select.text)
		warehouse_select.click()

		date_button = find_elements_by_presence('i.el-range__icon')[0].click()
		time.sleep(1)
		date_start = date.today() - timedelta(days=arrival_time_start)
		date_end = date.today() - timedelta(days=arrival_time_end)
		print(date_start, date_end)
		date_select(date_start, date_end)
		search_button = find_element_by_clickable('button.el-button.el-icon-search').click()
		time.sleep(2)
		download_button = find_element_by_clickable('button.el-button.el-icon-download').click()

		comfirmed_tag = find_element_by_presence('div.el-message-box__container')
		if comfirmed_tag.text == '导出成功！':
			comfirmed_button = find_element_by_clickable('button.el-button--default').click()
		time.sleep(2)
		for item in find_elements_by_presence('ul.sideBox li'):
			if '下载管理' in item.text:
				download_mangement = item.find_element_by_css_selector('a[href="#/download"]').click()
				break
		download = find_element_by_presence('tr.el-table__row:nth-child(1)')
		while True:
			download_status = download.find_element_by_css_selector('td:nth-child(3)')
			if download_status.text == '完成':
				download_access = download.find_element_by_css_selector('td:nth-child(5) a')
				download_url = download_access.get_attribute('href')
				browser.execute_script('window.open()')
				browser.switch_to.window(browser.window_handles[2])
				# ActionChains(browser).move_to_element(download_access).click(download_access).perform()
				browser.get(download_url)
				break
			time.sleep(2)
		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'C:\Users\aesdhj\Desktop'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('branch' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, '服务站支线{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(1)
		browser.close()
		browser.switch_to.window(browser.window_handles[1])
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
	browser.quit()
else:
	pass



sys.exit(0)
