import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None
import warnings
warnings.filterwarnings('ignore')
from tqdm import tqdm
import openpyxl
from openpyxl.styles import Alignment, Font
import sys
from appium import webdriver as appium_webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver as selenium_webdriver
import re
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta, date
import time
from selenium.webdriver import ActionChains
import shutil
import os


complaints_cms_action, tc_short_check_action = 0, 0
product_list_action, tuan_information_action = 0, 0
day_gap = 2
complaints_cms_start, complaints_cms_end = 3, 0
tc_short_check_start, tc_short_check_end = day_gap, day_gap
mapping = {
	'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
	'七': 7, '八': 8, '九': 9, '十': 10, '十一': 11, '十二': 12,
}
mapping_ltc = {
	'武进直配': '武进LTC', '无锡直配': '无锡LTC', '上海嘉定直配': '上海嘉定LTC', '安徽直配': '安徽LTC', '溧阳直配1': '溧阳LTC'
}


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


def get_verification_code_qb():
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
		if '阿里云' in verification_text:
			return verification_code
			break


# 导出CMS客诉记录_________________________________________________________________________________________________________
if complaints_cms_action == 1:
	print('导出CMS客诉记录')
	browser = selenium_webdriver.Chrome()
	browser.get('https://cms.nicetuan.net/admin/signin/')
	wait = WebDriverWait(browser, 30)
	user = find_element_by_presence('input#telphone').send_keys('')
	code_button = find_element_by_clickable('div.msgBtn').click()
	verification_code = get_verification_code()
	print(verification_code)
	code_confirmed = browser.switch_to.alert.accept()
	password = find_element_by_presence('input#captcha').send_keys(verification_code)
	login = find_element_by_clickable('button.submit').click()
	for items in [
		# ('hanzhou', 'div.dropdown-menu.open ul.dropdown-menu.inner > li:nth-child(3)'),
		('wenzhou', 'div.dropdown-menu.open ul.dropdown-menu.inner > li:nth-child(4)'),
		('jiangsu', 'div.dropdown-menu.open ul.dropdown-menu.inner > li:nth-child(6)'),
		('hefei', 'div.dropdown-menu.open ul.dropdown-menu.inner > li:nth-child(7)'),
	]:
		complaints_menu = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'i.fa-tasks')))
		if complaints_menu:
			menu_button = find_element_by_clickable('a.responsive-toggler').click()
		complaints_menu = find_element_by_clickable('i.fa-tasks').click()
		complaints_cms = find_element_by_clickable('li.open ul.sub-menu li:nth-child(2)').click()
		station_button = find_element_by_clickable('label#city_id_label').click()
		station_select = find_element_by_clickable(items[1])
		print(items[0], station_select.text)
		station_select.click()
		search_type = Select(find_element_by_presence('select#time_search')).select_by_value('suborderid')

		date_start = date.today() - timedelta(days=complaints_cms_start)
		date_end = date.today() - timedelta(days=complaints_cms_end)
		print(date_start, date_end)
		time_selects = find_elements_by_presence('div.datetimepicker-dropdown-bottom-right')

		start_time_button = find_element_by_clickable('input#starttime').click()
		time.sleep(1)
		start_month = time_selects[0].find_element_by_css_selector('div.datetimepicker-days th.switch')
		start_month = re.search('(.*?)月', start_month.text).group(1)
		start_month = mapping[start_month]
		print(start_month)
		pre_button = time_selects[0].find_element_by_css_selector('div.datetimepicker-days th.prev')
		next_button = time_selects[0].find_element_by_css_selector('div.datetimepicker-days th.next')
		time.sleep(1)
		if date_start.month < start_month:
			pre_button.click()
		elif date_start.month > start_month:
			next_button.click()
		elif date_start.month == start_month:
			pass
		item_today = time_selects[0].find_element_by_css_selector('div.datetimepicker-days td.day.active')
		if int(item_today.text) == date_start.day:
			item_today.click()
		else:
			for item in time_selects[0].find_elements_by_css_selector('div.datetimepicker-days td[class="day"]'):
				if int(item.text) == date_start.day:
					item.click()
					break
		for item in time_selects[0].find_elements_by_css_selector('div.datetimepicker-hours span.hour'):
			if item.text == '0:00':
				item.click()
				break
		for item in time_selects[0].find_elements_by_css_selector('div.datetimepicker-minutes span.minute'):
			if item.text == '0:00':
				item.click()
				break
		time.sleep(1)

		end_time_button = find_element_by_clickable('input#endtime').click()
		time.sleep(1)
		end_month = time_selects[1].find_element_by_css_selector('div.datetimepicker-days th.switch')
		end_month = re.search('(.*?)月', end_month.text).group(1)
		end_month = mapping[end_month]
		print(end_month)
		pre_button = time_selects[1].find_element_by_css_selector('div.datetimepicker-days th.prev')
		next_button = time_selects[1].find_element_by_css_selector('div.datetimepicker-days th.next')
		time.sleep(1)
		if date_end.month < end_month:
			pre_button.click()
		elif date_end.month > end_month:
			next_button.click()
		elif date_end.month == end_month:
			pass
		item_today = time_selects[1].find_element_by_css_selector('div.datetimepicker-days td.day.active')
		if int(item_today.text) == date_end.day:
			item_today.click()
		else:
			for item in time_selects[1].find_elements_by_css_selector('div.datetimepicker-days td[class="day"]'):
				if int(item.text) == date_end.day:
					item.click()
					break
		for item in time_selects[1].find_elements_by_css_selector('div.datetimepicker-hours span.hour'):
			if item.text == '0:00':
				item.click()
				break
		for item in time_selects[1].find_elements_by_css_selector('div.datetimepicker-minutes span.minute'):
			if item.text == '0:00':
				item.click()
				break
		time.sleep(1)
		search_button = find_element_by_clickable('button[value="search"]').click()
		while True:
			data_len = find_element_by_presence('form#search_form div.pull-left b').text
			if int(data_len) > 0:
				print(data_len)
				break
			time.sleep(2)
		time.sleep(1)
		download_button = find_element_by_clickable('button#export_btn').click()
		time.sleep(1)
		download_tag = find_element_by_clickable('div.btn.sure').click()
		time.sleep(1)
		download_confirmed_text = find_element_by_presence('div.myAlertBox > p').text
		download_confirmed_text = download_confirmed_text.split(':')[1]
		print(download_confirmed_text)
		time.sleep(1)
		download_confirmed = find_element_by_clickable('div.btn.sure').click()
		time.sleep(1)
		download_access = find_element_by_clickable('a[title="查看报表"]').click()

		while True:
			j = 0
			for i in range(1, 2):
				id_text = find_element_by_presence('table#adminlist > tbody > tr:nth-child({}) td.text-center:nth-child(1)'.format(i)).text
				download_condition = find_element_by_presence('table#adminlist > tbody > tr:nth-child({}) td.text-center:nth-child(4)'.format(i)).text
				if (id_text == download_confirmed_text) and (download_condition == '导出完成'):
					download_access = find_element_by_presence('table#adminlist > tbody > tr:nth-child({}) td.text-center:nth-child(7)'.format(i))
					ActionChains(browser).move_to_element(download_access).click(download_access).perform()
					j += 1
			if j > 0:
				break
			time.sleep(5)

		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'D:\文档\工作\十荟团\temp\司机团长差异'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('xlsx' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, 'complaints_cms_{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(2)
		back_button = find_element_by_clickable('ul.page-breadcrumb a[href="http://workorder-cms.nicetuan.net/"]').click()
	browser.quit()
else:
	pass
print('-'*50)


# 导出TC差异明细_________________________________________________________________________________________________________
if tc_short_check_action == 1:
	print('导出TC差异明细')
	browser = selenium_webdriver.Chrome()
	browser.get('https://wms.nicetuan.net/login')
	wait = WebDriverWait(browser, 20)
	user = find_element_by_presence('input.input_all.name').send_keys('13770995264')
	code_button = find_element_by_clickable('button#code').click()
	verification_code = get_verification_code()
	print(verification_code)
	password = find_element_by_presence('input.input_all_code.password').send_keys(verification_code)
	warehouse_select = find_element_by_clickable('a.textbox-icon.combo-arrow').click()
	warehouse = find_element_by_clickable('div#_easyui_combobox_i1_4').click()
	login = find_element_by_clickable('a#login.button.hide').click()
	#20210604
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
		tc_short_check = find_element_by_clickable('ul[role="menubar"] span:nth-child(8) > li li.el-menu-item:nth-child(4)').click()
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
			if data_table != '暂无数据':
				break
			time.sleep(1)
		downnload_button = find_elements_by_presence('button.el-icon-download')[1].click()
		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'D:\文档\工作\十荟团\temp\司机团长差异'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('服务站差异含明细' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, '服务站差异含明细_{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(1)
		time.sleep(1)
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
	browser.quit()
	time.sleep(1)
else:
	pass
print('-'*50)


# 导出出库商品列表——————————————————————————————————————————————————————————————————————————————————————————————————————————
if product_list_action == 1:
	print('导出商品列表')
	browser = selenium_webdriver.Chrome()
	browser.get('https://wms.nicetuan.net/login')
	wait = WebDriverWait(browser, 20)
	user = find_element_by_presence('input.input_all.name').send_keys('13770995264')
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
	lvyuezhongxin = find_element_by_clickable('div#_easyui_tree_19 span.tree-title').click()
	browser.switch_to.window(browser.window_handles[1])
	lvyueguanli = find_element_by_clickable('ul[role="menubar"] span:nth-child(1) > li').click()

	for items in [
		# ('wenzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(3)'),
		# ('hanzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(2)'),
		('jiangsu', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(1)'),
		# ('xuzhou', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(5)'),
		# ('hefei', 'div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(6)'),
	]:
		select_rows = find_elements_by_presence('div.el-row')

		select_rows_2 = select_rows[1]
		station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
		time.sleep(1)
		station_select = find_element_by_clickable(items[1])
		print(station_select.text)
		station_select.click()
		if items[0] == 'jiangsu':
			select_rows_1 = select_rows[0]
			time.sleep(1)
			date_button = select_rows_1.find_elements_by_css_selector('div.el-col')[0].click()
			time.sleep(1)
			date_start = date.today() - timedelta(days=day_gap)
			date_end = date.today() - timedelta(days=day_gap)
			print(date_start, date_end)
			date_select(date_start, date_end)

			deliver_condition_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[3].click()
			delivered_conditions = find_elements_by_presence('div[x-placement="bottom-start"] li.el-select-dropdown__item')
			for item in delivered_conditions:
				if item.text == '已发货' or item.text == '未发货':
					item.click()
				time.sleep(0.5)
		select_rows_6 = select_rows[5]
		search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
		time.sleep(1)
		while True:
			loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
			if loading_tag:
				break
		select_rows_7 = select_rows[6]
		exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
		exported_button = find_element_by_presence('button.el-button--success')
		ActionChains(browser).move_to_element(exported_button).click(exported_button).perform()

		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'D:\文档\工作\十荟团\temp\司机团长差异'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('出库商品明细列表' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, '出库商品明细列表_{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(3)

		back_button = find_element_by_clickable('div.align-right button:nth-child(1)').click()
		time.sleep(1)
	browser.quit()
	time.sleep(1)
else:
	pass
print('-'*50)


# 导出团长信息—————————————————————————————————————————————————————————————————————————————————————————————————————————
if tuan_information_action == 1:
	print('导出团长信息')
	option = selenium_webdriver.ChromeOptions()
	option.add_experimental_option(
		'excludeSwitches',
		['enable-automation']
	)
	browser = selenium_webdriver.Chrome(chrome_options=option)
	browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
		"source": """
			Object.defineProperty(navigator, 'webdriver', {
				get: () => undefined
			})
			"""
	})

	browser.get('https://das.base.shuju.aliyun.com/product/BI_SHT.htm?menuId=6ggopgm2p3g')
	wait = WebDriverWait(browser, 30)
	if '账号密码登录' in browser.page_source:
		tab_elemetns = find_elements_by_presence('div.tabs-item-text')
		for item in tab_elemetns:
			if item.text == '账号密码登录':
				item.click()
	else:
		frame_tab = find_element_by_presence('iframe#login-iframe-2019')
		browser.switch_to.frame(frame_tab)
		user_password_tab = find_elements_by_presence('div.ability-tabs-item-text')[1]
		user_password_tab.click()

	user_password_frame = find_element_by_presence('iframe#alibaba-login-box')
	browser.switch_to.frame(user_password_frame)
	user = find_element_by_presence('input[name="fm-login-id"]').send_keys('dingtalk_laurnm')
	password = find_element_by_presence('input[name="fm-login-password"]').send_keys('103728aesdhj')
	login = find_element_by_presence('div.fm-btn button').click()
	time.sleep(1)

	def is_element_present(item):
		from selenium.common.exceptions import NoSuchElementException
		try:
			element = browser.find_element_by_css_selector(item)
		except NoSuchElementException as e:
			return False
		else:
			return True

	if is_element_present('iframe#baxia-dialog-content'):
		slider_frame = find_element_by_presence('iframe#baxia-dialog-content')
		browser.switch_to.frame(slider_frame)
		slider_area = find_element_by_presence('div#nc_1__scale_text')
		slider = find_element_by_presence('span#nc_1_n1z')
		ActionChains(browser).drag_and_drop_by_offset(slider, slider_area.size['width'], slider_area.size['height']).perform()

		# user_password_frame = find_element_by_presence('iframe#alibaba-login-box')
		# browser.switch_to.frame(user_password_frame)

	i = 0
	while True:
		if '你正在使用手机短信验证身份' not in browser.page_source:
			i += 1
			time.sleep(2)
			if i >= 3:
				break
		else:
			code_button = find_element_by_clickable('button#J_GetCode').click()
			verification_code_qb = get_verification_code_qb()
			print(verification_code_qb)
			password = find_element_by_presence('div.checkcode-warp input#J_Checkcode').send_keys(
				verification_code_qb)
			submit_button = find_element_by_clickable('div.submit button').click()
			break

	for item in find_elements_by_presence('div.top-nav-text'):
		if '团长' in item.text:
			item.click()
			break
	for item in find_elements_by_presence('span.menu-item-title'):
		if '团长基础信息表' in item.text:
			item.click()
			break

	data_iframe = find_element_by_presence('iframe.portal-iframe-container')
	browser.switch_to.frame(data_iframe)
	i = 0
	while True:
		if '数据加载中' not in browser.page_source:
			i += 1
			time.sleep(2)
			if i >= 3:
				break
		else:
			time.sleep(3)

	query_field = find_element_by_presence('div.query-component span.label-name')
	query_field.click()
	options_button = find_element_by_presence('i.qbi-scope-antd-dropdown-trigger').click()
	time.sleep(1)
	options_items = find_elements_by_presence('li.qbi-scope-antd-dropdown-menu-item')
	for item in options_items:
		if '创建取数' in item.text:
			item.click()
			break
	time.sleep(1)
	for item in find_elements_by_presence('div.qbi-scope-antd-modal-footer button'):
		if '确 定' in item.text:
			item.click()
			break
	time.sleep(1)
	while True:
		loading_tag = find_element_by_presence('tbody.qbi-scope-antd-table-tbody > tr:nth-child(1) td:nth-child(3)')
		if '成功' in loading_tag.text:
			loading_access = find_element_by_presence('tbody.qbi-scope-antd-table-tbody > tr:nth-child(1) td:nth-child(4) i')
			datetime_tag = find_element_by_presence('tbody.qbi-scope-antd-table-tbody > tr:nth-child(1) td:nth-child(2)')
			print(datetime_tag.text)
			ActionChains(browser).move_to_element(loading_access).click(loading_access).perform()
			break
		time.sleep(1)
		refresh_button = find_elements_by_presence('button.qbi-scope-antd-btn-primary')[1].click()
		time.sleep(10)
	download_path = r'C:\Users\aesdhj\Downloads'
	destination_path = r'D:\文档\工作\十荟团\temp\司机团长差异'
	while True:
		i = 0
		file_list = os.listdir(download_path)
		for file in file_list:
			if ('新交叉表' in file) and ('crdownload' not in file):
				download_file_path = os.path.join(download_path, file)
				destination_file_path = os.path.join(destination_path, '{}.xlsx'.format('tuan_information'))
				shutil.move(download_file_path, destination_file_path)
				i += 1
		if i > 0:
			break
		time.sleep(5)
	print('-'*50)
	browser.quit()
	time.sleep(1)
else:
	pass


# 计算_________________________________________________________________________________________________________
print('开始计算')
path = r'D:\文档\工作\十荟团\temp\司机团长差异'
if complaints_cms_action == 1:
	complaints_cms = pd.read_csv(r'D:\文档\工作\十荟团\temp\complaints_csm_14days.csv')
	complaints_cms['下单时间'] = pd.to_datetime(complaints_cms['下单时间'])
	complaints_cms_old = complaints_cms[complaints_cms['下单时间'] >= (date.today() - timedelta(days=31))]
	file_list = os.listdir(path)
	file_df = []
	for file_nm in tqdm(file_list):
		if ('complaints_cms' in file_nm) and file_nm != '每日':
			temp_df = pd.read_excel(os.path.join(path, file_nm))
			file_df.append(temp_df)
	complaints_cms = pd.concat(file_df, axis=0, ignore_index=True)
	complaints_cms['下单时间'] = pd.to_datetime(complaints_cms['下单时间'])
	print(complaints_cms['下单时间'].min(), complaints_cms['下单时间'].max())
	complaints_cms = pd.concat([complaints_cms, complaints_cms_old], axis=0, ignore_index=True)
	complaints_cms = complaints_cms.drop_duplicates().reset_index(drop=True)
	complaints_cms.to_csv(r'D:\文档\工作\十荟团\temp\complaints_csm_14days.csv', index=False, encoding='utf_8_sig')
	complaints_cms = complaints_cms.sort_values(['下单时间'], ascending=False)
	tuan_id_to_name = complaints_cms[['团长ID', '团长姓名']].drop_duplicates(['团长ID'], keep='first').reset_index(drop=True)
	product_id_to_spcification = complaints_cms[['商品ID', '规格ID']].drop_duplicates(['商品ID'], keep='first').reset_index(drop=True)
else:
	tuan_information = pd.read_excel(r'D:\文档\工作\十荟团\temp\司机团长差异\tuan_information.xlsx', usecols=['团长id', '团长姓名'])
	tuan_information = tuan_information.rename(columns={'团长id': '团长ID'})
	tuan_id_to_name = tuan_information[['团长ID', '团长姓名']].drop_duplicates(['团长ID']).reset_index(drop=True)

	file_list = os.listdir(path)
	file_df = []
	for file_nm in tqdm(file_list):
		if ('出库商品明细列表' in file_nm) and file_nm != '每日':
			temp_df = pd.read_excel(os.path.join(path, file_nm))
			file_df.append(temp_df)
	product_list = pd.concat(file_df, axis=0, ignore_index=True)
	product_list['商品简称'] = product_list['商品简称'].fillna(method='ffill')
	product_list['商品ID'] = product_list.apply(lambda x: re.search('（(\d+)）', x['商品简称']).group(1), axis=1)
	product_list['商品ID'] = product_list['商品ID'].astype(int)
	product_id_to_spcification = product_list[['商品ID', '规格ID']].drop_duplicates(['商品ID']).reset_index(drop=True)
	product_id_to_spcification['规格ID'] = product_id_to_spcification['规格ID'].astype(int)
	product_id_to_spcification['规格ID'] = product_id_to_spcification['规格ID'].astype(str)

# 2021-09-01
df = pd.read_excel(r'D:\文档\工作\十荟团\temp\EXPORT_COMPLAINTS-DETAIL.xlsx')
df = df[df['大区'] == '华东']
df = df.rename(columns={'子订单id': '子订单Id', 'TC服务站名称': 'TC服务站', '子站名称': '站点名称', '送货日期/快递时间': '快递时间', '一级分类': '后端分类', '小区名称': '配送地址'})
df['更新时间'] = df.apply(lambda x: x['更新时间'] if pd.isna(x['更新时间']) else str(x['更新时间'])[:4] + '-' + str(x['更新时间'])[4:6] + '-' + str(x['更新时间'])[6:8], axis=1)
df['快递时间'] = df.apply(lambda x: x['快递时间'] if pd.isna(x['快递时间']) else str(int(x['快递时间']))[:4] + '-' + str(int(x['快递时间']))[4:6] + '-' + str(int(x['快递时间']))[6:8], axis=1)
df['订单时间'] = df.apply(lambda x: x['订单时间'] if pd.isna(x['订单时间']) else str(x['订单时间'])[:4] + '-' + str(x['订单时间'])[4:6] + '-' + str(x['订单时间'])[6:8], axis=1)
complaints_detail = df
# complaints_detail = pd.read_csv(r'D:\文档\工作\十荟团\temp\EXPORT_COMPLAINTS-DETAIL.csv')

# 20210612
if '司机' in list(complaints_detail.columns):
	complaints_detail = complaints_detail.rename(columns={'司机': '司机名称'})

complaints_detail['快递时间'] = pd.to_datetime(complaints_detail['快递时间'])
complaints_detail = complaints_detail[complaints_detail['快递时间'] == (date.today() - timedelta(days=day_gap))]
complaints_detail_unknown = complaints_detail[pd.isna(complaints_detail['TC服务站'])]
print('无服务站记录:', complaints_detail_unknown.shape)
complaints_detail = complaints_detail[~pd.isna(complaints_detail['TC服务站'])]
complaints_detail = complaints_detail[
	(complaints_detail['客诉原因'] == '缺货退款') |
	(complaints_detail['客诉原因'] == '仓库差异') |
	(complaints_detail['客诉原因'] == 'TC仓差异') |
	(complaints_detail['客诉原因'] == '运输差异') |
	(complaints_detail['客诉原因'] == '供应商未到货') |
	(complaints_detail['客诉原因'] == '团长丢件') |
	(complaints_detail['客诉原因'] == '团长发错货') |
	(complaints_detail['客诉原因'] == '供应商来货错误') |
	(complaints_detail['客诉原因'] == '仓库配错商品') |
	(complaints_detail['客诉原因'] == '司机送错商品') |
	(complaints_detail['客诉原因'] == 'TC仓分拣错误')
	]
complaints_tc_product = complaints_detail.groupby(['快递时间', '站点名称', 'TC服务站', '司机名称', '团长id', '配送地址', '商品id', '商品名称', '后端分类'], as_index=False)['售后数量'].sum()
complaints_tc_product['TC服务站'] = complaints_tc_product.apply(lambda x: mapping_ltc[x['TC服务站']] if x['TC服务站'] in mapping_ltc else x['TC服务站'], axis=1)
# 20210525更新，TC名称变化
complaints_tc_product['TC服务站'] = complaints_tc_product.apply(lambda x: re.sub('\d+', '', x['TC服务站']), axis=1)
complaints_tc_product = complaints_tc_product.rename(columns={'商品id': '商品ID', '团长id': '团长ID', '站点名称': '子站'})
complaints_tc_product = complaints_tc_product.merge(product_id_to_spcification, on='商品ID', how='left')
complaints_tc_product = complaints_tc_product.merge(tuan_id_to_name, on='团长ID', how='left')
temp_son_to_father_old = pd.read_csv(r'D:\文档\工作\十荟团\temp\temp_son_to_father_old.csv')
complaints_tc_product = complaints_tc_product.merge(temp_son_to_father_old, on='子站', how='left')
# 2021-08-06
# complaints_tc_product = complaints_tc_product[complaints_tc_product['主站名称'] != '杭州市']
print(complaints_tc_product.shape, complaints_tc_product['售后数量'].sum())
# complaints_tc_product['规格ID'] = complaints_tc_product['规格ID'].astype(int)
# complaints_tc_product['规格ID'] = complaints_tc_product['规格ID'].astype(str)
# print(complaints_tc_product.head())

file_list = os.listdir(path)
file_df = []
for file_nm in tqdm(file_list):
	if ('服务站差异含明细' in file_nm) and file_nm != '每日':
		temp_df = pd.read_excel(os.path.join(path, file_nm))
		file_df.append(temp_df)
tc_difference = pd.concat(file_df, axis=0, ignore_index=True)
tc_difference['送货日期'] = pd.to_datetime(tc_difference['送货日期'])
print(tc_difference['送货日期'].min(), tc_difference['送货日期'].max())
tc_difference['少货数量'] = tc_difference['少货数量'].fillna(0)
tc_difference['少货数量'] = tc_difference['少货数量'].astype(int)
tc_difference['sku或规格Id'] = tc_difference['sku或规格Id'].astype(str)
tc_difference = tc_difference.rename(columns={'sku或规格Id': '规格ID', '服务站': 'TC服务站'})
tc_difference['TC服务站'] = tc_difference.apply(lambda x: mapping_ltc[x['TC服务站']] if x['TC服务站'] in mapping_ltc else x['TC服务站'], axis=1)
# 20210525更新，TC名称变化
tc_difference['TC服务站'] = tc_difference.apply(lambda x: re.sub('\d+', '', x['TC服务站']), axis=1)
tc_difference_no_tuan = tc_difference[pd.isna(tc_difference['团长姓名'])]
tc_difference_no_tuan = tc_difference_no_tuan.rename(columns={'少货数量': '少货数量_无团长姓名'})
tc_difference_no_tuan = tc_difference_no_tuan.groupby(['TC服务站', '规格ID'], as_index=False)['少货数量_无团长姓名'].sum()
tc_difference = tc_difference[~pd.isna(tc_difference['团长姓名'])]
tc_difference = tc_difference.groupby(['TC服务站', '团长姓名', '规格ID'], as_index=False)['少货数量'].sum()
complaints_tc_product = complaints_tc_product.merge(tc_difference_no_tuan, on=['TC服务站', '规格ID'], how='left')
complaints_tc_product['少货数量_无团长姓名'] = complaints_tc_product['少货数量_无团长姓名'].fillna(0)
complaints_tc_product = complaints_tc_product.merge(tc_difference, on=['TC服务站', '团长姓名', '规格ID'], how='left')
complaints_tc_product['少货数量'] = complaints_tc_product['少货数量'].fillna(0)

short_supply_detail = pd.read_csv(r'D:\文档\工作\十荟团\temp\缺货改期明细统计.csv', encoding='gbk')
short_supply_detail['日期'] = pd.to_datetime(short_supply_detail['日期'])
short_supply_detail = short_supply_detail[short_supply_detail['日期'] == (date.today() - timedelta(days=day_gap))]
short_supply_detail['缺货件数'] = short_supply_detail['缺货件数'].fillna(0)
short_supply_detail['缺货件数'] = short_supply_detail['缺货件数'].astype(int)
short_supply_detail_ltc = short_supply_detail[short_supply_detail['来源'] != 'TC']
short_supply_detail_ltc = short_supply_detail_ltc.rename(columns={'规格id': '规格ID', '来源': 'TC服务站', '缺货件数': '主仓缺货数量_ltc'})
short_supply_detail_ltc = short_supply_detail_ltc.groupby(['TC服务站', '规格ID'], as_index=False)['主仓缺货数量_ltc'].sum()
complaints_tc_product = complaints_tc_product.merge(short_supply_detail_ltc, on=['TC服务站', '规格ID'], how='left')
complaints_tc_product['主仓缺货数量_ltc'] = complaints_tc_product['主仓缺货数量_ltc'].fillna(0)
print('主仓缺货数量_ltc', complaints_tc_product['主仓缺货数量_ltc'].sum())
short_supply_detail = short_supply_detail[short_supply_detail['来源'] == 'TC']
short_supply_detail['主站名称'] = short_supply_detail['城市仓'].map({'杭州仓': '杭州市', '温州仓': '浙南十荟团', '溧阳仓': '江苏十荟团', '合肥仓': '安徽十荟团', '徐州仓': '徐州十荟团'})
short_supply_detail = short_supply_detail.groupby(['主站名称', '规格id'], as_index=False)['缺货件数'].sum()
short_supply_detail = short_supply_detail.rename(columns={'规格id': '规格ID', '缺货件数': '主仓缺货数量'})
complaints_tc_product = complaints_tc_product.merge(short_supply_detail, on=['主站名称', '规格ID'], how='left')
complaints_tc_product['主仓缺货数量'] = complaints_tc_product['主仓缺货数量'].fillna(0)
complaints_tc_product['主仓缺货数量'] = complaints_tc_product.apply(lambda x: 0 if x['TC服务站'] in ['安徽LTC', '杭州市区-直配', '杭州市区直配', '杭州-嘉兴直配', '新北LTC', '武进LTC', '无锡LTC', '溧阳LTC', '上海嘉定LTC', '虚拟TC'] else x['主仓缺货数量'], axis=1)
complaints_tc_product['差异'] = complaints_tc_product['售后数量'] - complaints_tc_product['少货数量_无团长姓名'] - complaints_tc_product['少货数量'] - complaints_tc_product['主仓缺货数量'] - complaints_tc_product['主仓缺货数量_ltc']

complaints_tc_product = complaints_tc_product[complaints_tc_product['差异'] > 0]
complaints_tc_product['快递时间'] = complaints_tc_product['快递时间'].dt.date
complaints_tc_product = complaints_tc_product[~complaints_tc_product['TC服务站'].isin(['杭州-嘉兴LTC', '杭州市区LTC', '杭州市区直配', '杭州-嘉兴直配', '杭州市区-直配', '新北LTC', '溧阳LTC', '上海嘉定LTC', '虚拟TC'])]
complaints_tc_product = complaints_tc_product[['快递时间', '主站名称', 'TC服务站', '司机名称', '团长ID', '团长姓名', '配送地址', '后端分类', '商品ID', '规格ID', '商品名称', '差异']]
complaints_tc_product = complaints_tc_product.sort_values(['快递时间', '主站名称', 'TC服务站', '司机名称', '团长ID', '团长姓名', '配送地址', '后端分类', '商品ID', '规格ID'])
# complaints_tc_product['规格ID'] = complaints_tc_product['规格ID'].astype(int)
print(complaints_tc_product[(pd.isna(complaints_tc_product['团长姓名'])) | (pd.isna(complaints_tc_product['规格ID']))]['差异'].sum(), complaints_tc_product['差异'].sum())
path = 'D:/文档/工作/十荟团/temp/司机团长差异/每日/{}_司机团长差异明细.xlsx'.format((date.today() - timedelta(days=day_gap)).strftime('%Y%m%d'))
complaints_tc_product.to_excel(path, sheet_name='明细', index=False)

wb = openpyxl.load_workbook(path)
sheet = wb.get_active_sheet()
font = Font(name='微软雅黑', size=11, bold=False)
for row in sheet.rows:
	for cell in row:
		cell.font = font
sheet.column_dimensions['A'].width = 12
sheet.column_dimensions['B'].width = 11
sheet.column_dimensions['C'].width = 20
sheet.column_dimensions['D'].width = 16
sheet.column_dimensions['E'].width = 13
sheet.column_dimensions['K'].width = 20
sheet.freeze_panes = 'A2'
wb.save(path)

path = r'D:\文档\工作\十荟团\temp\司机团长差异\每日'
file_list_14days = ['{}_司机团长差异明细.xlsx'.format((datetime.today() - timedelta(days=i)).strftime('%Y%m%d')) for i in range(2, 16)]
file_list = os.listdir(path)
file_df = []
for file_nm in tqdm(file_list):
	if file_nm in file_list_14days:
		temp_df = pd.read_excel(os.path.join(path, file_nm))
		file_df.append(temp_df)
tuan_dirver_difference = pd.concat(file_df, axis=0, ignore_index=True)
tuan_dirver_difference['快递时间'] = pd.to_datetime(tuan_dirver_difference['快递时间'])
tuan_dirver_difference['日期'] = tuan_dirver_difference['快递时间']
tuan_dirver_difference = tuan_dirver_difference.rename(columns={'后端分类': '分类', '规格id': '规格ID'})
tuan_dirver_difference['规格ID'] = tuan_dirver_difference.apply(lambda x: x['规格ID'] if pd.isna(x['规格ID']) else str(int(x['规格ID'])), axis=1)
tuan_dirver_difference.to_csv(r'D:\文档\工作\十荟团\temp\temp_tuan_driver_difference.csv', index=False, encoding='utf_8_sig')

sys.exit(0)