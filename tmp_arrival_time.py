import pandas as pd
import numpy as np
pd.options.display.max_rows = None
pd.options.display.max_columns = None
from datetime import datetime, timedelta, date
import re
import warnings
warnings.filterwarnings('ignore')
from appium import webdriver as appium_webdriver
from selenium import webdriver as selenium_webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import time
import sys
import os
import shutil


dispatch_start, dispatch_end = 1, 0
dispatch_action = 1


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


# 导出准点率数据----------------------------------------------------------------------------------------------------------
if dispatch_action == 1:
	for items in [
		# ('hanzhou', 'div#_easyui_combobox_i1_2'),
		# ('wenzhou', 'div#_easyui_combobox_i1_3'),
		('jiangsu', 'div#_easyui_combobox_i1_4'),
		# ('hefei', 'div#_easyui_combobox_i1_8'),
		# ('xuzhou', 'div#_easyui_combobox_i1_5'),
	]:
		warehouse_list = []
		warehouse_list.append(items[0])
		browser = selenium_webdriver.Chrome()
		browser.get('https://wms.nicetuan.net/login')
		wait = WebDriverWait(browser, 20)
		user = find_element_by_presence('input.input_all.name').send_keys('')
		code_button = find_element_by_clickable('button#code').click()
		verification_code = get_verification_code()
		print(verification_code)
		password = find_element_by_presence('input.input_all_code.password').send_keys(verification_code)
		time.sleep(1)
		warehouse_select = find_element_by_clickable('a.textbox-icon.combo-arrow').click()
		warehouse = find_element_by_clickable(items[1])
		warehouse_list.append(warehouse.text)
		warehouse.click()
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

		tms_center = find_element_by_clickable('div#_easyui_tree_20 span.tree-title').click()
		browser.switch_to.window(browser.window_handles[1])
		for item in find_elements_by_presence('ul.sideBox li'):
			if '报表管理' in item.text:
				item.click()
				time.sleep(1)
				dispatch_task_report = item.find_element_by_css_selector('ul[role="menu"] li:nth-child(1)').click()
				time.sleep(1)
				break
		browser.close()
		browser.switch_to.window(browser.window_handles[0])
		tms_center = find_element_by_clickable('div#_easyui_tree_20 span.tree-title').click()
		browser.switch_to.window(browser.window_handles[1])
		for item in find_elements_by_presence('ul.sideBox li'):
			if '报表管理' in item.text:
				item.click()
				time.sleep(1)
				dispatch_task_report = item.find_element_by_css_selector('ul[role="menu"] li:nth-child(1)').click()
				time.sleep(1)
				break
		warehouse_button = find_element_by_clickable('div.is-justify-space-between > div:nth-child(1) span.el-input__suffix').click()
		warehouse_select = find_element_by_clickable('div[x-placement="bottom-start"] li.selected').click()
		warehouse_list.append(find_element_by_clickable('div[x-placement="bottom-start"] li.selected').text)

		date_start = date.today() - timedelta(days=dispatch_start)
		date_end = date.today() - timedelta(days=dispatch_end)
		print(date_start, date_end)

		date_button = find_element_by_clickable('div.el-date-editor--daterange').click()
		time.sleep(1)
		date_select(date_start, date_end)
		search_button = find_element_by_clickable('button.el-button--primary').click()
		while True:
			loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
			if loading_tag == True:
				break
		while True:
			download_button = find_element_by_clickable('button.el-button--success').click()
			comfirmed_tag = find_element_by_presence('div.el-message-box__container')
			comfirmed_button = find_element_by_clickable('button.el-button--default').click()
			if comfirmed_tag.text == '导出成功！':
				break
			time.sleep(2)
		print(warehouse_list)
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
				browser.get(download_url)
				# ActionChains(browser).move_to_element(download_access).click(download_access).perform()
				break
			time.sleep(5)
		download_path = r'C:\Users\aesdhj\Downloads'
		destination_path = r'D:\文档\工作\十荟团\temp\准点率'
		while True:
			i = 0
			file_list = os.listdir(download_path)
			for file in file_list:
				if ('dispatch_task_report' in file) and ('crdownload' not in file):
					download_file_path = os.path.join(download_path, file)
					destination_file_path = os.path.join(destination_path, 'dispatch_task_report_{}.xlsx'.format(items[0]))
					shutil.move(download_file_path, destination_file_path)
					i += 1
			if i > 0:
				break
			time.sleep(5)
		time.sleep(1)
		browser.quit()
else:
	pass

# 计算-------------------------------------------------------------------------------------------------------------------
cities_jzh = [
	'淮北市', '丽水市', '铜陵市', '安庆市', '马鞍山市', '蚌埠市', '嘉兴市', '宿州市', '台州市', '池州市', '南京市', '徐州市', '六安市',
	'合肥市', '滁州市', '湖州市', '阜阳市', '镇江市', '衢州市', '淮南市', '黄山市', '扬州市', '南通市', '金华市', '无锡市', '温州市',
	'连云港市', '苏州市', '盐城市', '淮安市', '宁波市', '泰州市', '绍兴市', '亳州市', '宣城市', '芜湖市', '常州市', '宿迁市', '舟山市',
	'杭州市', '上海市', '商丘市']
tuan_id_adjust = pd.read_csv(r'D:\文档\工作\十荟团\temp\tuan_id_adjust.csv', encoding='gbk')
id_city = {}
for item in tuan_id_adjust.values:
	id_city[item[0]] = item[1]

dispatch_task_report_jiangsu = pd.read_excel(r'D:\文档\工作\十荟团\temp\准点率\dispatch_task_report_jiangsu.xlsx')
dispatch_task_report_jiangsu['主站名称'] = '江苏十荟团'
dispatch_task_report_hanzhou = pd.read_excel(r'D:\文档\工作\十荟团\temp\准点率\dispatch_task_report_hanzhou.xlsx')
dispatch_task_report_hanzhou['主站名称'] = '杭州市'
dispatch_task_report_wenzhou = pd.read_excel(r'D:\文档\工作\十荟团\temp\准点率\dispatch_task_report_wenzhou.xlsx')
dispatch_task_report_wenzhou['主站名称'] = '浙南十荟团'
dispatch_task_report_hefei = pd.read_excel(r'D:\文档\工作\十荟团\temp\准点率\dispatch_task_report_hefei.xlsx')
dispatch_task_report_hefei['主站名称'] = '安徽十荟团'
dispatch_task_report_xuzhou = pd.read_excel(r'D:\文档\工作\十荟团\temp\准点率\dispatch_task_report_xuzhou.xlsx')
dispatch_task_report_xuzhou['主站名称'] = '徐州十荟团'
dispatch_task_report = pd.concat([dispatch_task_report_jiangsu, dispatch_task_report_hanzhou, dispatch_task_report_wenzhou, dispatch_task_report_hefei, dispatch_task_report_xuzhou], axis=0, ignore_index=True)
# 20210525更新，TC名称变化
dispatch_task_report['配送方式'] = dispatch_task_report.apply(lambda x: re.sub('\d+', '', x['配送方式']), axis=1)

dispatch_task_report['地级市'] = dispatch_task_report.apply(lambda x:
					re.search('省(.*?市)', x['团长地址']).group(1) if re.search('省(.*?市)', x['团长地址']) != None
					else (re.search('(.*?市)', x['团长地址']).group(1) if re.search('(.*?市)', x['团长地址']) != None
					else None), axis=1)
dispatch_task_report['地级市'] = dispatch_task_report.apply(lambda x: x['地级市'] if x['地级市'] in cities_jzh else (id_city[x['团长id']] if x['团长id'] in id_city else None), axis=1)
temp = dispatch_task_report[~dispatch_task_report['地级市'].isin(cities_jzh)]
temp = temp.drop_duplicates(['团长id'])
if len(temp) > 0:
	temp.to_csv(r'C:\Users\aesdhj\Desktop\temp.csv', index=False, encoding='utf_8_sig')
	print('id_to_city:', len(temp))
	sys.exit(0)

dispatch_task_report['团长地址'] = dispatch_task_report.apply(lambda x: x['团长地址'].replace('\n', ''), axis=1)
dispatch_task_report['配送完成时间'] = pd.to_datetime(dispatch_task_report['配送完成时间'])
dispatch_task_report['送货日期'] = pd.to_datetime(dispatch_task_report['送货日期'])
dispatch_task_report['团点数'] = dispatch_task_report.apply(lambda x: 1 if x['团长商品数'] > 0 else 0, axis=1)
dispatch_task_report['hour'] = dispatch_task_report['配送完成时间'].dt.hour

dispatch_task_report['12点准点团点数'] = dispatch_task_report.apply(lambda x: 1 if (x['团长商品数'] > 0) * (x['hour'] < 12) * (x['配送状态'] == '已完成') else 0, axis=1)
dispatch_task_report['14点准点团点数'] = dispatch_task_report.apply(lambda x: 1 if (x['团长商品数'] > 0) * (x['hour'] < 14) * (x['配送状态'] == '已完成') else 0, axis=1)

mapping = {
	'南京淮阴MTC': '南京淮阴区TC', '南京楚州MTC': '南京楚州区TC', '南京秦淮STC': '南京秦淮TC', '象山丹阳MTC': '象山丹西MTC',
	'杭州-江干钱江TC': '杭州-江干下沙TC', '瑞安东山STC': '瑞安东山MTC', '安徽合肥肥西县TC': '安徽合肥市肥西县TC',
	'安徽合肥市小庙TC': '安徽合肥小庙镇TC', '安徽合肥南区TC': '安徽合肥市南区TC', '安徽合肥东区TC': '安徽合肥市东区TC', '安徽合肥肥东区TC': '安徽合肥市东区TC',
	'杭州市区直配': '杭州市区LTC', '杭州-嘉兴直配': '杭州-嘉兴LTC', '苏州甪直MTC': '苏州甪直TC', '杭州市区-直配': '杭州市区LTC',
	'武进直配': '武进LTC', '无锡直配': '无锡LTC', '上海嘉定直配': '上海嘉定LTC', '安徽直配': '安徽LTC', '溧阳直配1': '溧阳LTC',
	'溧阳直配': '溧阳LTC'
}

dispatch_task_report['准点商品数'] = dispatch_task_report['团长商品数'] * dispatch_task_report['14点准点团点数']
dispatch_task_report['准点货值'] = dispatch_task_report['团长货值'] * dispatch_task_report['14点准点团点数']
# dispatch_task_report['作业区'] = dispatch_task_report.apply(lambda x: '直配' if x['作业区'] == '未配置' else ('直配' if x['配送方式'] in ['杭州市区直配', '杭州-嘉兴直配', '新北LTC', '武进LTC', '无锡LTC', '溧阳LTC', '上海嘉定LTC', '虚拟TC'] else x['作业区']), axis=1)
dispatch_task_report['作业区'] = dispatch_task_report.apply(lambda x: '直配' if x['配送方式'] in ['安徽LTC', '杭州-嘉兴LTC', '杭州市区LTC', '杭州市区直配', '杭州市区-直配', '杭州-嘉兴直配', '新北LTC', '武进LTC', '无锡LTC', '溧阳LTC', '上海嘉定LTC', '虚拟TC'] else 'TC', axis=1)
dispatch_task_report['配送方式'] = dispatch_task_report.apply(lambda x: '溧阳LTC' if (x['主站名称'] == '江苏十荟团') * (x['配送方式'] == '直配') else x['配送方式'], axis=1)
dispatch_task_report['配送方式'] = dispatch_task_report.apply(lambda x: '浙南虚拟TC' if (x['主站名称'] == '浙南十荟团') * (x['配送方式'] == '虚拟TC') else x['配送方式'], axis=1)
dispatch_task_report['配送方式'] = dispatch_task_report.apply(lambda x: '安徽LTC' if (x['主站名称'] == '安徽十荟团') * (x['配送方式'] == '直配') else x['配送方式'], axis=1)

tc_address_lon_lat = pd.read_csv(r'D:\文档\工作\十荟团\temp\tc_address_lon_lat.csv')
tc_lost = []
for tc in list(set(dispatch_task_report['配送方式'])):
	if tc not in list(set(tc_address_lon_lat['配送方式'])):
		tc_lost.append(tc)
if len(tc_lost) > 0:
	print('tc_lost:', tc_lost)
	# sys.exit(0)
temp = dispatch_task_report.groupby(['主站名称', '送货日期'], as_index=False)['团点数', '14点准点团点数'].sum()
temp['14点准点率'] = temp['14点准点团点数'] / temp['团点数']
print(temp)

dispatch_task_report_past = pd.read_csv(r'D:\文档\工作\十荟团\temp\temp_dispatch_task_report.csv')
dispatch_task_report_past['送货日期'] = pd.to_datetime(dispatch_task_report_past['送货日期'])
dispatch_task_report_past = dispatch_task_report_past[(dispatch_task_report_past['送货日期'] != (date.today() - timedelta(days=16))) & (dispatch_task_report_past['送货日期'] != (date.today() - timedelta(days=1)))]

dispatch_task_report = pd.concat([dispatch_task_report, dispatch_task_report_past], axis=0, ignore_index=True)
dispatch_task_report['配送方式'] = dispatch_task_report.apply(lambda x: mapping[x['配送方式']] if x['配送方式'] in mapping else x['配送方式'], axis=1)

dispatch_task_report.to_csv(r'D:\文档\工作\十荟团\temp\temp_dispatch_task_report.csv', index=False, encoding='utf_8_sig')

sys.exit(0)