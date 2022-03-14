import pandas as pd
import os
from tqdm import tqdm
import numpy as np
from datetime import datetime, timedelta, date
import sys
import warnings
from appium import webdriver as appium_webdriver
from selenium import webdriver as selenium_webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import re
import time
import shutil

warnings.filterwarnings('ignore')


sales_start, sales_end = 1, 1
complaints_start, complaints_end = 14, 1
partner_product_action = 0
complaint_details_action = 0
partner_product_action_qb = 1
complaint_details_action_qb = 1


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


def find_element_by_clickable(item):
	element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, item)))
	return element


def find_element_by_presence(item):
	element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, item)))
	return element


def find_elements_by_presence(item):
	elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, item)))
	return elements


def login_bi():
	user = find_element_by_presence('input[placeholder="请输入已注册手机号"]').send_keys('')
	code_button = find_element_by_clickable('div.el-input-group__append').click()
	verification_code = get_verification_code()
	print(verification_code)
	password = find_element_by_presence('input[placeholder="请输入验证码"]').send_keys(verification_code)
	login = find_element_by_clickable('button.login-submit').click()


def date_select(date_start, date_end):
	date_panel_left_head_text = find_element_by_presence(
		'div.el-date-range-picker__content.is-left div.el-date-range-picker__header').text
	date_panel_left_head_text_month = int(re.search('(.*?) 年 (.*?) 月', date_panel_left_head_text).group(2))
	print(date_panel_left_head_text_month)
	date_panel_left_forward = find_element_by_presence(
		'div.el-date-range-picker__content.is-left div.el-date-range-picker__header button.el-icon-arrow-left')
	date_panel_left_back = find_element_by_presence(
		'div.el-date-range-picker__content.is-left div.el-date-range-picker__header button.el-icon-arrow-right')

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
	elif date_start.month > date_panel_left_head_text_month:
		date_panel_left_table = find_elements_by_presence(
			'div.el-date-range-picker__content.is-left table td.next-month')
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
		date_panel_left_back.click()
	else:
		print('date_end_error')
		sys.exit(0)

# 导出团长售后明细—————————————————————————————————————————————————————————————————————————————————————————————————————————
if partner_product_action_qb == 1:
	print('导出团长商品销售明细')
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

		user_password_frame = find_element_by_presence('iframe#alibaba-login-box')
		browser.switch_to.frame(user_password_frame)
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
		if '团长商品销售日报_团期' in item.text:
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

	query_field = find_element_by_presence('div.query-field-wrapper.horizontal-label span.label-name')
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
		refresh_button = find_elements_by_presence('button.qbi-scope-antd-btn-primary')[1].click()
		time.sleep(10)
	download_path = r'C:\Users\aesdhj\Downloads'
	destination_path = r'D:\文档\工作\十荟团\temp\团长商品销量报表'
	while True:
		i = 0
		file_list = os.listdir(download_path)
		for file in file_list:
			if ('新交叉表' in file) and ('crdownload' not in file):
				download_file_path = os.path.join(download_path, file)
				destination_file_path = os.path.join(destination_path, '{}.xlsx'.format(date.today() - timedelta(days=1)))
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

if partner_product_action == 1:
	print('导出团长售后明细')
	browser = selenium_webdriver.Chrome()
	browser.get('http://bi.nicetuan.net/#/login')
	wait = WebDriverWait(browser, 30)
	login_bi()
	operation_data_center = find_element_by_clickable('img[src="../../../../static/img/bg/sys-sht.png"]').click()
	operation_data_menu = find_element_by_clickable('i.yhdx-icon-yunyingshuju').click()
	partner_product_sales = find_element_by_clickable('li.is-opened ul[role="menu"] li:nth-child(4)').click()
	station_button = find_element_by_clickable('div.sebox div.sht-date-range:nth-child(3)').click()
	station_selects = find_elements_by_presence('div[style="display: flex;"] div.pop-text')
	for item in station_selects:
		item.click()
	time.sleep(1)
	station_comfirmed = find_element_by_clickable('div.popbottom button').click()
	date_button = find_element_by_clickable('div.sebox div.sht-date-range:nth-child(4)').click()

	date_end = date.today() - timedelta(days=sales_end)
	date_start = date.today() - timedelta(days=sales_start)
	print(date_start, date_end)
	time.sleep(1)
	date_select(date_start, date_end)

	search_button = find_element_by_clickable('div.sebox div.el-row  > div:nth-child(7)').click()

	# 等待返回
	while True:
		data_len_text = find_element_by_presence('div.el-pagination.is-background').text
		data_len_text = int(re.search('共 (.*?) 条', data_len_text).group(1))
		if data_len_text > 0:
			break
		time.sleep(3)

	download_button = find_element_by_clickable('div.sebox div.el-row  > div:nth-child(8)').click()
	download_success = find_element_by_presence('p.el-message__content')
	print(download_success.text)
	time.sleep(15)

	while True:
		download_access = find_element_by_clickable('div.download-wrapper').click()
		my_download = find_element_by_presence('span.el-dialog__title')
		download = find_element_by_presence('div.el-dialog__body div.el-table__body-wrapper tr.el-table__row:nth-child(1)')
		result = download.find_element_by_css_selector('td:nth-child(4)')
		if result.text == '处理完成':
			download_out = download.find_element_by_css_selector('td.is-hidden span')
			ActionChains(browser).move_to_element(download_out).click(download_out).perform()
			close = find_elements_by_presence('button.el-dialog__headerbtn')[1]
			time.sleep(1)
			close.click()
			break
		close = find_elements_by_presence('button.el-dialog__headerbtn')[1]
		time.sleep(1)
		close.click()
		time.sleep(15)

	download_path = r'C:\Users\aesdhj\Downloads'
	destination_path = r'D:\文档\工作\十荟团\temp\团长商品销量报表'
	while True:
		i = 0
		file_list = os.listdir(download_path)
		for file in file_list:
			if ('EXPORT_PARTNER-PRODUCT-DATA' in file) and ('crdownload' not in file):
				download_file_path = os.path.join(download_path, file)
				destination_file_path = os.path.join(destination_path, '{}.csv'.format(date.today() - timedelta(days=sales_start)))
				shutil.move(download_file_path, destination_file_path)
				i += 1
		if i > 0:
			break
		time.sleep(15)
	print('-'*50)
	browser.quit()
	time.sleep(1)
else:
	pass

# 导出售后工单—————————————————————————————————————————————————————————————————————————————————————————————————————————
if complaint_details_action == 1:
	print('导出售后工单')
	browser = selenium_webdriver.Chrome()
	browser.get('http://bi.nicetuan.net/#/login')
	wait = WebDriverWait(browser, 20)
	login_bi()
	operation_data_center = find_element_by_clickable('img[src="../../../../static/img/bg/sys-sht.png"]').click()
	complaints_menu = find_element_by_clickable('i.yhdx-icon-kehulvyue').click()
	complaint_details = find_element_by_clickable('li.is-opened ul[role="menu"] li:nth-child(1)').click()
	date_button = find_element_by_clickable('div.sebox div.sht-date-range:nth-child(2)').click()

	date_end = date.today() - timedelta(days=complaints_end)
	date_start = date.today() - timedelta(days=complaints_start)
	print(date_start, date_end)
	time.sleep(1)
	date_select(date_start, date_end)

	station_button = find_element_by_clickable('div.sebox div.sht-date-range:nth-child(3)').click()
	station_selects = find_elements_by_presence('div[style="display: flex;"] div.pop-text')
	for item in station_selects:
		item.click()
	time.sleep(1)
	station_comfirmed = find_element_by_clickable('div.popbottom button').click()
	time.sleep(1)
	buttons = find_elements_by_presence('div.sebox div.el-row div')
	for item in buttons:
		if '询' in item.text:
			item.click()
	while True:
		data_len_text = find_element_by_presence('div.el-pagination.is-background').text
		data_len_text = int(re.search('共 (.*?) 条', data_len_text).group(1))
		if data_len_text > 0:
			break
		time.sleep(3)

	for item in buttons:
		if '导出' in item.text:
			item.click()
	download_success = find_element_by_presence('p.el-message__content')
	print(download_success.text)
	time.sleep(10)

	while True:
		download_access = find_element_by_clickable('div.download-wrapper').click()
		my_download = find_element_by_presence('span.el-dialog__title')
		download = find_element_by_presence('div.el-dialog__body div.el-table__body-wrapper tr.el-table__row:nth-child(1)')
		result = download.find_element_by_css_selector('td:nth-child(4)')
		if result.text == '处理完成':
			while True:
				download_out = download.find_element_by_css_selector('td.is-hidden span')
				ActionChains(browser).move_to_element(download_out).click(download_out).perform()
				download_num = download.find_element_by_css_selector('td:nth-child(6)')
				if int(download_num.text) > 0:
					break
				time.sleep(1)
			close = find_elements_by_presence('button.el-dialog__headerbtn')[1]
			time.sleep(1)
			close.click()
			break
		close = find_elements_by_presence('button.el-dialog__headerbtn')[1]
		time.sleep(1)
		close.click()
		time.sleep(10)

	download_path = r'C:\Users\aesdhj\Downloads'
	destination_path = r'D:\文档\工作\十荟团\temp'
	while True:
		i = 0
		file_list = os.listdir(download_path)
		for file in file_list:
			if ('EXPORT_COMPLAINTS-DETAIL' in file) and ('crdownload' not in file):
				download_file_path = os.path.join(download_path, file)
				destination_file_path = os.path.join(destination_path, '{}.csv'.format('EXPORT_COMPLAINTS-DETAIL'))
				shutil.move(download_file_path, destination_file_path)
				i += 1
		if i > 0:
			break
		time.sleep(10)
	print('-'*50)
	time.sleep(3)
	browser.quit()
else:
	pass

# 导出售后工单—————————————————————————————————————————————————————————————————————————————————————————————————————————
if complaint_details_action_qb == 1:
	print('导出售后明细')
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
	browser.maximize_window()
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
		if '体验' in item.text:
			item.click()
			break
	for item in find_elements_by_presence('span.menu-item-title'):
		if '客诉工单明细表' in item.text:
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

	search_button = find_element_by_clickable('i.icon-qbi-shaixuan').click()
	time.sleep(1)
	date_start_button = find_elements_by_presence('span.date-picker')[0].click()
	time.sleep(1)
	month_text = find_element_by_presence('a.qbi-scope-antd-calendar-month-select')
	month_text = re.search('(.*?)月', month_text.text).group(1)
	month_text = int(month_text)
	date_start_x = datetime.today() - timedelta(days=14)
	month_x = date_start_x.month
	date_start_x = f'{date_start_x.year}年{date_start_x.month}月{date_start_x.day}日'
	print(date_start_x)
	if month_text != month_x:
		find_element_by_presence('a.qbi-scope-antd-calendar-prev-month-btn').click()
	time.sleep(1)
	for item in find_elements_by_presence('table.qbi-scope-antd-calendar-table td'):
		date_text = item.get_attribute('title')
		if date_text == date_start_x:
			item.click()
			break
	i = 0
	while True:
		if '数据加载中' not in browser.page_source:
			i += 1
			time.sleep(2)
			if i >= 3:
				break
		else:
			time.sleep(3)

	query_field = find_element_by_presence('div.query-field-wrapper.horizontal-label span.label-name')
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
		refresh_button = find_elements_by_presence('button.qbi-scope-antd-btn-primary')[1].click()
		time.sleep(10)
	download_path = r'C:\Users\aesdhj\Downloads'
	destination_path = r'D:\文档\工作\十荟团\temp'
	while True:
		i = 0
		file_list = os.listdir(download_path)
		for file in file_list:
			if ('新交叉表' in file) and ('crdownload' not in file):
				download_file_path = os.path.join(download_path, file)
				destination_file_path = os.path.join(destination_path, '{}.xlsx'.format('EXPORT_COMPLAINTS-DETAIL'))
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
# 计算——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
print('开始计算')
temp_partner_product_past = pd.read_csv(r'D:\文档\工作\十荟团\temp\temp_partner_product.csv')
temp_partner_product_past['日期'] = pd.to_datetime(temp_partner_product_past['日期'])
temp_partner_product_past = temp_partner_product_past[temp_partner_product_past['日期'] != (date.today() - timedelta(days=15))]
if len(temp_partner_product_past[temp_partner_product_past['日期'] == (date.today() - timedelta(days=1))]) > 0:
	print('已有昨天数据')
	print('日期', len(list(set(temp_partner_product_past['日期']))))
	temp_partner_product = temp_partner_product_past
else:
	print('无昨天数据')
	# file_list_______________________________________________________________________________________________________
	file_list_35days = [(datetime.today() - timedelta(days=i)).strftime('%Y-%m-%d')+'.xlsx' for i in range(1, 2)]
	# partner_product__________________________________________________________________________________________________
	path = r'D:\文档\工作\十荟团\temp\团长商品销量报表'
	file_list = os.listdir(path)
	file_df = []
	for file_nm in tqdm(file_list):
		if file_nm in file_list_35days:
			temp_df = pd.read_excel(os.path.join(path, file_nm))
			temp_df['日期'] = file_nm.split('.')[0]
			file_df.append(temp_df)
	partner_product_today = pd.concat(file_df, axis=0, ignore_index=True)
	partner_product_today['日期'] = pd.to_datetime(partner_product_today['日期'])
	partner_product_today = partner_product_today.rename(columns={'主站': '主站名称'})
	print(partner_product_today.shape)
	partner_product_today = partner_product_today[partner_product_today['主站名称'].isin(['安徽十荟团', '杭州市', '江苏十荟团', '徐州十荟团', '浙南十荟团'])]
	print(partner_product_today.shape)
	# partner_product_today['商品管理分类名'] = partner_product_today['商品管理分类名'].fillna('未分类')
	# partner_product_today = partner_product_today.rename(columns={'商品管理分类名': '后端分类'})
	print(partner_product_today['子订单数'].sum(), partner_product_today['nmv'].sum())
	# temp_partner_product_today = partner_product_today.groupby(['主站名称', '子站', '日期', '后端分类'], as_index=False)['子订单数', 'nmv'].sum()
	temp_partner_product_today = partner_product_today.groupby(['主站名称', '子站', '日期'], as_index=False)['子订单数', 'nmv'].sum()
	# print(temp_partner_product_today['子订单数'].sum(), temp_partner_product_today['nmv'].sum())
	temp_partner_product = pd.concat([temp_partner_product_today, temp_partner_product_past], axis=0, ignore_index=True)
	print('日期', len(list(set(temp_partner_product['日期']))))

temp = temp_partner_product[['主站名称', '子站']].drop_duplicates().reset_index(drop=True)
temp_old = pd.read_csv(r'D:\文档\工作\十荟团\temp\temp_son_to_father_old.csv')
temp = pd.concat([temp, temp_old], axis=0, ignore_index=True)
temp = temp.drop_duplicates(['子站'], keep='first').reset_index(drop=True)
temp.to_csv(r'D:\文档\工作\十荟团\temp\temp_son_to_father_old.csv', index=False, encoding='utf_8_sig')
mapping = {}
for item in temp.values:
	mapping[item[1]] = item[0]

# complaints_detail_____________________________________________________________________________________________________
# 2021-09-01
df = pd.read_excel(r'D:\文档\工作\十荟团\temp\EXPORT_COMPLAINTS-DETAIL.xlsx')
df = df[df['大区'] == '华东']
df = df.rename(columns={'子订单id': '子订单Id', 'TC服务站名称': 'TC服务站', '子站名称': '站点名称', '送货日期/快递时间': '快递时间', '一级分类': '后端分类'})
df['更新时间'] = df.apply(lambda x: x['更新时间'] if pd.isna(x['更新时间']) else str(x['更新时间'])[:4] + '-' + str(x['更新时间'])[4:6] + '-' + str(x['更新时间'])[6:8], axis=1)
df['快递时间'] = df.apply(lambda x: x['快递时间'] if pd.isna(x['快递时间']) else str(int(x['快递时间']))[:4] + '-' + str(int(x['快递时间']))[4:6] + '-' + str(int(x['快递时间']))[6:8], axis=1)
df['订单时间'] = df.apply(lambda x: x['订单时间'] if pd.isna(x['订单时间']) else str(x['订单时间'])[:4] + '-' + str(x['订单时间'])[4:6] + '-' + str(x['订单时间'])[6:8], axis=1)
complaints_detail = df

# complaints_detail = pd.read_csv(r'D:\文档\工作\十荟团\temp\EXPORT_COMPLAINTS-DETAIL.csv')
# 2021-08-16
if len(complaints_detail) == len(complaints_detail[pd.isna(complaints_detail['责任归属'])]):
	print('无责任归属')
	complaints_detail = complaints_detail.drop('责任归属', axis=1)
	file_list_35days = [(datetime.today() - timedelta(days=i)).strftime('%Y-%m-%d') + '.xlsx' for i in range(1, 15)]
	path = r'D:\文档\工作\十荟团\temp\责任归属'
	file_list = os.listdir(path)
	file_df = []
	for file_nm in tqdm(file_list):
		if file_nm in file_list_35days:
			temp_df = pd.read_excel(os.path.join(path, file_nm))
			file_df.append(temp_df)
	id_res = pd.concat(file_df, axis=0, ignore_index=True)
	id_res = id_res.rename(columns={'子订单id': '子订单Id', 'accountabilitys': '责任归属'})
	id_res['子订单Id'] = id_res.apply(lambda x: re.sub('D', '', x['子订单Id']), axis=1)
	id_res['子订单Id'] = id_res['子订单Id'].astype(np.int64)
	id_res = id_res[['子订单Id', '责任归属']]
	id_res = id_res.drop_duplicates('子订单Id')
	complaints_detail = complaints_detail.merge(id_res, on='子订单Id', how='left')

complaints_detail = complaints_detail.rename(columns={'子订单Id': '子订单id'})
complaints_detail['更新时间'] = pd.to_datetime(complaints_detail['更新时间'])
# 2021-09-01
start = (date.today() - timedelta(days=14))
end = (date.today() - timedelta(days=1))
complaints_detail = complaints_detail[complaints_detail['更新时间'] >= start]
complaints_detail = complaints_detail[complaints_detail['更新时间'] <= end]

complaints_detail['订单时间'] = pd.to_datetime(complaints_detail['订单时间'])
complaints_detail['子订单id'] = complaints_detail['子订单id'].astype(str)
complaints_detail = complaints_detail[complaints_detail['责任归属'] != '团长']
complaints_detail = complaints_detail[complaints_detail['责任归属'] != '用户']
complaints_detail_cancel_not_follow = complaints_detail[((pd.isna(complaints_detail['责任归属'])) & (complaints_detail['工单类型'] == '撤单'))]
print('撤单未跟踪:', len(complaints_detail_cancel_not_follow))
complaints_detail = complaints_detail[~((pd.isna(complaints_detail['责任归属'])) & (complaints_detail['工单类型'] == '撤单'))]

# 加上秒赔的售后金额
miaopei = complaints_detail[complaints_detail['赔付类型'] == '秒赔']
complaints_detail = complaints_detail[complaints_detail['赔付类型'] != '秒赔']
complaints_detail_ids = list(set(complaints_detail['子订单id']))
miaopei_not_follow = miaopei[~miaopei['子订单id'].isin(complaints_detail_ids)]
print('miaopei_not_follow:', len(miaopei_not_follow))
miaopei = miaopei[miaopei['子订单id'].isin(complaints_detail_ids)]
miaopei = miaopei.rename(columns={'售后金额': '售后金额+'})
miaopei = miaopei[['更新时间', '子订单id', '售后金额+']]
print('售后', miaopei['售后金额+'].sum())
temp = complaints_detail[['更新时间', '子订单id', '售后金额']].copy()
temp = temp.drop_duplicates(['更新时间', '子订单id'])
miaopei = miaopei.merge(temp, on=['更新时间', '子订单id'], how='left')
miaopei_next = miaopei[pd.isna(miaopei['售后金额'])]
miaopei = miaopei[~pd.isna(miaopei['售后金额'])]
miaopei = miaopei.groupby(['更新时间', '子订单id'], as_index=False)['售后金额+'].sum()
miaopei_next = miaopei_next.drop('售后金额', axis=1)
miaopei_next = miaopei_next.rename(columns={'售后金额+': '售后金额++'})
miaopei_next = miaopei_next.groupby('子订单id', as_index=False)['售后金额++'].sum()
complaints_detail = complaints_detail.merge(miaopei, on=['更新时间', '子订单id'], how='left')
complaints_detail['售后金额+'] = complaints_detail['售后金额+'].fillna(0)
complaints_detail = complaints_detail.merge(miaopei_next, on='子订单id', how='left')
complaints_detail['售后金额++'] = complaints_detail['售后金额++'].fillna(0)
print('售后', complaints_detail['售后金额+'].sum()+complaints_detail['售后金额++'].sum())
complaints_detail['售后金额'] = complaints_detail['售后金额'] + complaints_detail['售后金额+'] + complaints_detail['售后金额++']
complaints_detail = complaints_detail.drop(['售后金额+', '售后金额++'], axis=1)

# 2021-09-01
temp = complaints_detail[pd.isna(complaints_detail['责任归属'])]
print('责任为空', temp['售后数量'].sum(), temp['售后金额'].sum())
complaints_detail = complaints_detail[~pd.isna(complaints_detail['责任归属'])]

complaints_detail = complaints_detail.sort_values(['更新时间', '子订单id'])
complaints_detail['子订单id_next'] = complaints_detail['子订单id'].shift(1)
complaints_detail['售后子订单数'] = complaints_detail.apply(lambda x: 1 if x['子订单id'] != x['子订单id_next'] else 0, axis=1)

complaints_detail['责任归属'] = complaints_detail['责任归属'].fillna('未跟踪')
complaints_detail['客诉原因'] = complaints_detail['客诉原因'].fillna('未跟踪')
complaints_detail['后端分类'] = complaints_detail['后端分类'].fillna('未分类')
check_list = list(set(complaints_detail['站点名称']))
for item in check_list:
	if item not in mapping:
		print('站点名称', item)
complaints_detail['主站名称'] = complaints_detail['站点名称'].map(mapping)
complaints_detail = complaints_detail.rename(columns={'更新时间': '日期', '站点名称': '子站'})
print(complaints_detail['售后子订单数'].sum())
print(complaints_detail['售后金额'].sum())
temp_complaints_detail = complaints_detail.groupby(['主站名称', '子站', '日期', '后端分类', '工单类型', '客诉原因', '责任归属'], as_index=False)['售后子订单数', '售后金额', '售后数量'].sum()
print(temp_complaints_detail['售后子订单数'].sum())
print(temp_complaints_detail['售后金额'].sum())

# 字段扩展________________________________________________________________________________________________________________
temp_partner_product_only = temp_partner_product[['主站名称', '子站', '日期']].drop_duplicates()
temp_complaints_detail_only = temp_complaints_detail[['后端分类', '工单类型', '客诉原因', '责任归属']].drop_duplicates()
partner_product_part = np.repeat(temp_partner_product_only.values, temp_complaints_detail_only.shape[0], axis=0)
complaints_detail_part = np.tile(temp_complaints_detail_only.values, (temp_partner_product_only.shape[0], 1))
part = np.concatenate([partner_product_part, complaints_detail_part], axis=1)
part_df = pd.DataFrame(part, columns=list(temp_partner_product_only.columns) + list(temp_complaints_detail_only.columns))
part_df['售后子订单数'] = 0
part_df['售后金额'] = 0
part_df['售后数量'] = 0
temp_complaints_detail = pd.concat([temp_complaints_detail, part_df], axis=0, ignore_index=True)
temp_complaints_detail = temp_complaints_detail.groupby(['主站名称', '子站', '日期', '后端分类', '工单类型', '客诉原因', '责任归属'], as_index=False)['售后子订单数', '售后金额', '售后数量'].sum()

# 2021-05-18更新, 2021-08-06
# temp_complaints_detail = temp_complaints_detail[temp_complaints_detail['主站名称'] != '杭州市']
# temp_partner_product = temp_partner_product[temp_partner_product['主站名称'] != '杭州市']
temp_complaints_detail.to_csv(r'D:\文档\工作\十荟团\temp\temp_complaints_detail.csv', index=False, encoding='utf_8_sig')
temp_partner_product.to_csv(r'D:\文档\工作\十荟团\temp\temp_partner_product.csv', index=False, encoding='utf_8_sig')

# top_______________________________________________________________________________________________________________
top_detail = complaints_detail.groupby(['主站名称', '子站', '日期', '商品id', '商品名称', '后端分类', '客诉原因', '责任归属', '团长id'], as_index=False)['售后子订单数', '售后金额', '售后数量'].sum()
top_detail_reason = top_detail.groupby(['主站名称', '日期', '商品id', '商品名称', '客诉原因'], as_index=False)['售后子订单数', '售后金额'].sum()
top_detail_reason['售后子订单数'] = round(top_detail_reason['售后子订单数'])
top_detail_reason['售后金额'] = round(top_detail_reason['售后金额'])
top_detail_reason_q = top_detail_reason.sort_values(['主站名称', '日期', '商品id', '商品名称', '售后子订单数'], ascending=False)
top_detail_reason_q = top_detail_reason_q.drop_duplicates(['主站名称', '日期', '商品id', '商品名称'])
top_detail_reason_q['主要客诉原因_q'] = top_detail_reason_q['客诉原因'] + '/' + top_detail_reason_q['售后子订单数'].astype(str)
top_detail_reason_q = top_detail_reason_q[['主站名称', '日期', '商品id', '商品名称', '主要客诉原因_q']]
top_detail = top_detail.merge(top_detail_reason_q, on=['主站名称', '日期', '商品id', '商品名称'], how='left')
top_detail_reason_m = top_detail_reason.sort_values(['主站名称', '日期', '商品id', '商品名称', '售后金额'], ascending=False)
top_detail_reason_m = top_detail_reason_m.drop_duplicates(['主站名称', '日期', '商品id', '商品名称'])
top_detail_reason_m['主要客诉原因_m'] = top_detail_reason_m['客诉原因'] + '/' + top_detail_reason_m['售后金额'].astype(str)
top_detail_reason_m = top_detail_reason_m[['主站名称', '日期', '商品id', '商品名称', '主要客诉原因_m']]
top_detail = top_detail.merge(top_detail_reason_m, on=['主站名称', '日期', '商品id', '商品名称'], how='left')
tuan_information = pd.read_excel(r'D:\文档\工作\十荟团\temp\司机团长差异\tuan_information.xlsx', usecols=['团长id', '团长姓名'])
tuan_id_to_name = tuan_information[['团长id', '团长姓名']].drop_duplicates(['团长id']).reset_index(drop=True)
top_detail = top_detail.merge(tuan_id_to_name, on='团长id', how='left')
# 2021-05-18更新, 2021-08-06
# top_detail = top_detail[top_detail['主站名称'] != '杭州市']
top_detail.to_csv(r'D:\文档\工作\十荟团\temp\temp_top_detail.csv', index=False, encoding='utf_8_sig')

sys.exit(0)