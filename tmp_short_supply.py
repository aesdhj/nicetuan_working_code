import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None
import numpy as np
from datetime import datetime, timedelta, date
import warnings
warnings.filterwarnings('ignore')
import sys
from appium import webdriver as appium_webdriver
from selenium import webdriver as selenium_webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
import pandas as pd
import warnings
warnings.filterwarnings('ignore')


lvyue_start, lvyue_end = 0, 0
data_list = []
total_shipments_action = 1


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


# 导出出库总件数——————————————————————————————————————————————————————————————————————————————————————————————————————————
if total_shipments_action == 1:
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
	lvyuezhongxin = find_element_by_clickable('div#_easyui_tree_19 span.tree-title').click()

	browser.switch_to.window(browser.window_handles[1])
	lvyueguanli = find_element_by_clickable('ul[role="menubar"] span:nth-child(1) > li').click()

	select_rows = find_elements_by_presence('div.el-row')
	select_rows_1 = select_rows[0]
	time.sleep(1)
	date_button = select_rows_1.find_elements_by_css_selector('div.el-col')[0].click()
	time.sleep(1)

	date_start = date.today() - timedelta(days=lvyue_start)
	date_end = date.today() - timedelta(days=lvyue_end)
	print(date_start, date_end)
	date_select(date_start, date_end)

	select_rows_2 = select_rows[1]
	station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
	time.sleep(1)
	station_select = find_element_by_clickable('div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(1)')
	print(station_select.text)
	station_select.click()
	deliver_condition_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[3].click()
	delivered_conditions = find_elements_by_presence('div[x-placement="bottom-start"] li.el-select-dropdown__item')
	for item in delivered_conditions:
		if item.text == '已发货' or item.text == '未发货':
			item.click()
	deliver_condition_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[3].click()
	select_rows_6 = select_rows[5]
	search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
	while True:
		loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
		if loading_tag == True:
			break
	select_rows_7 = select_rows[6]
	exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
	exported_amount = find_element_by_presence('div.title')
	exported_amount_text = re.search('规格出库总数量 (.*?)，', exported_amount.text).group(1)
	print(exported_amount_text)
	data_list.append(['溧阳仓', exported_amount_text])
	time.sleep(1)

	# back_button = find_element_by_clickable('div.align-right button:nth-child(1)').click()
	# select_rows = find_elements_by_presence('div.el-row')
	# select_rows_2 = select_rows[1]
	# station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
	# time.sleep(1)
	# station_select = find_element_by_clickable('div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(3)')
	# print(station_select.text)
	# station_select.click()
	# select_rows_6 = select_rows[5]
	# search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
	# while True:
	# 	loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
	# 	if loading_tag == True:
	# 		break
	# select_rows_7 = select_rows[6]
	# exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
	# exported_amount = find_element_by_presence('div.title')
	# exported_amount_text = re.search('规格出库总数量 (.*?)，', exported_amount.text).group(1)
	# print(exported_amount_text)
	# data_list.append(['温州仓', exported_amount_text])
	# time.sleep(1)
	#
	# back_button = find_element_by_clickable('div.align-right button:nth-child(1)').click()
	# select_rows = find_elements_by_presence('div.el-row')
	# select_rows_2 = select_rows[1]
	# station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
	# time.sleep(1)
	# station_select = find_element_by_clickable('div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(4)')
	# print(station_select.text)
	# station_select.click()
	# select_rows_6 = select_rows[5]
	# search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
	# while True:
	# 	loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
	# 	if loading_tag == True:
	# 		break
	# select_rows_7 = select_rows[6]
	# exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
	# exported_amount = find_element_by_presence('div.title')
	# exported_amount_text = re.search('规格出库总数量 (.*?)，', exported_amount.text).group(1)
	# print(exported_amount_text)
	# data_list.append(['溧阳仓', exported_amount_text])
	# time.sleep(1)
	#
	# back_button = find_element_by_clickable('div.align-right button:nth-child(1)').click()
	# select_rows = find_elements_by_presence('div.el-row')
	# select_rows_2 = select_rows[1]
	# station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
	# time.sleep(1)
	# station_select = find_element_by_clickable('div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(5)')
	# print(station_select.text)
	# station_select.click()
	# select_rows_6 = select_rows[5]
	# search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
	# while True:
	# 	loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
	# 	if loading_tag == True:
	# 		break
	# select_rows_7 = select_rows[6]
	# exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
	# exported_amount = find_element_by_presence('div.title')
	# exported_amount_text = re.search('规格出库总数量 (.*?)，', exported_amount.text).group(1)
	# print(exported_amount_text)
	# data_list.append(['徐州仓', exported_amount_text])
	# time.sleep(1)
	#
	# back_button = find_element_by_clickable('div.align-right button:nth-child(1)').click()
	# select_rows = find_elements_by_presence('div.el-row')
	# select_rows_2 = select_rows[1]
	# station_button = select_rows_2.find_elements_by_css_selector('div.el-select--small')[0].click()
	# time.sleep(1)
	# station_select = find_element_by_clickable('div[x-placement="bottom-start"] li.el-select-dropdown__item:nth-child(6)')
	# print(station_select.text)
	# station_select.click()
	# select_rows_6 = select_rows[5]
	# search_button = select_rows_6.find_element_by_css_selector('i.el-icon-search').click()
	# while True:
	# 	loading_tag = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, 'div.el-loading-spinner')))
	# 	if loading_tag == True:
	# 		break
	# select_rows_7 = select_rows[6]
	# exported_product = select_rows_7.find_elements_by_css_selector('button')[0].click()
	# exported_amount = find_element_by_presence('div.title')
	# exported_amount_text = re.search('规格出库总数量 (.*?)，', exported_amount.text).group(1)
	# print(exported_amount_text)
	# data_list.append(['合肥仓', exported_amount_text])
	# time.sleep(1)
	browser.quit()

	df = pd.DataFrame(data_list, columns=['城市仓', '出库总件数'])
	df['日期'] = date.today()
	df['出库总件数'] = df['出库总件数'].astype(float)
	total_shipments = pd.read_csv(r'D:\文档\工作\十荟团\temp\出库总件数统计.csv')
	total_shipments = pd.concat([total_shipments, df], axis=0, ignore_index=True)
	total_shipments = total_shipments[['日期', '城市仓', '出库总件数']]
	total_shipments.to_csv(r'D:\文档\工作\十荟团\temp\出库总件数统计.csv', index=False, encoding='utf_8_sig')
else:
	pass

# 计算——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
total_shipments = pd.read_csv(r'D:\文档\工作\十荟团\temp\出库总件数统计.csv')
total_shipments['出库总件数'] = total_shipments['出库总件数'].fillna(0)
total_shipments['日期'] = pd.to_datetime(total_shipments['日期'])
total_shipments = total_shipments[(total_shipments['日期'] <= (date.today() - timedelta(days=0))) & (total_shipments['日期'] >= (date.today() - timedelta(days=14)))]
print(total_shipments['出库总件数'].sum())
mapping = {'南京仓': '南京市', '无锡仓': '江苏十荟团', '溧阳仓': '江苏十荟团', '温州仓': '浙南十荟团', '杭州仓': '杭州市', '合肥仓': '安徽十荟团', '徐州仓': '徐州十荟团'}
check_list = list(set(total_shipments['城市仓']))
for item in check_list:
	if item not in mapping:
		print(item)
total_shipments['主站名称'] = total_shipments['城市仓'].map(mapping)
total_shipments = total_shipments[['主站名称', '日期', '出库总件数']]
total_shipments = total_shipments.groupby(['主站名称', '日期'], as_index=False)['出库总件数'].sum()
print(total_shipments['出库总件数'].sum())
total_shipments.to_csv(r'D:\文档\工作\十荟团\temp\temp_total_shipments.csv', index=False, encoding='utf_8_sig')

short_supply_detail = pd.read_csv(r'D:\文档\工作\十荟团\temp\缺货改期明细统计.csv', encoding='gbk')
short_supply_detail['责任方'] = short_supply_detail.apply(lambda x: x['责任方'] if pd.isna(x['跟踪责任方']) else x['跟踪责任方'], axis=1)
short_supply_detail['缺货原因'] = short_supply_detail.apply(lambda x: x['缺货原因'] if pd.isna(x['跟踪缺货原因']) else x['跟踪缺货原因'], axis=1)
short_supply_detail = short_supply_detail.drop(['跟踪责任方', '跟踪缺货原因'], axis=1)
short_supply_detail['缺货件数'] = short_supply_detail['缺货件数'].fillna(0)
short_supply_detail['日期'] = pd.to_datetime(short_supply_detail['日期'])
short_supply_detail = short_supply_detail[(short_supply_detail['日期'] <= (date.today() - timedelta(days=0))) & (short_supply_detail['日期'] >= (date.today() - timedelta(days=14)))]
print(short_supply_detail['缺货件数'].sum())
short_supply_detail['缺货原因'] = short_supply_detail['缺货原因'].fillna('未标记')
short_supply_detail['采购'] = short_supply_detail['采购'].fillna('未标记')
short_supply_detail['责任方'] = short_supply_detail['责任方'].fillna('未标记')
temp = short_supply_detail.groupby('缺货原因', as_index=False)['缺货件数'].sum()
temp = temp.sort_values('缺货原因', ascending=False)
temp = temp[temp['缺货件数'] >= 100]
check_list = list(set(temp['缺货原因']))
mapping = {
	'未到货': '供应商未到货', '未到': '供应商未到货', '来货少': '供应商到货不足', '没到货': '供应商未到货', '供应商到货不足': '到货不足', '到货少': '供应商到货不足',
	'未配': '仓库未分拣', '没来货': '供应商未到货', '配货漏发': '仓库发错货', '品控拒收': '商品规格不符', '投线错误，还未收回': '仓库发错货',}
mapping_liyangcang = {
	'采购下错单': '采购下错单', '仓库转残品': '仓库转残品', '系统库存不准': '系统库存不准', '供应商未到货': '供应商未到货', '采购超卖': '采购超卖',
	'商品规格不符': '商品规格不符', '仓库发错货': '仓库发错货', '供应商到货不足': '供应商到货不足', '商品包装率不符': '商品包装率不符',}
mapping.update(mapping_liyangcang)
for item in check_list:
	if item not in mapping:
		print('缺货原因', item)
short_supply_detail['缺货原因'] = short_supply_detail.apply(lambda x: mapping[x['缺货原因']] if x['缺货原因'] in mapping else x['缺货原因'], axis=1)
short_supply_detail['责任方'] = short_supply_detail['责任方'].fillna('其他')
mapping = {'南京仓': '南京市', '无锡仓': '江苏十荟团', '溧阳仓': '江苏十荟团', '温州仓': '浙南十荟团', '杭州仓': '杭州市', '合肥仓': '安徽十荟团', '温州': '浙南十荟团', '徐州仓': '徐州十荟团'}
check_list = list(set(short_supply_detail['城市仓']))
for item in check_list:
	if item not in mapping:
		print('城市仓', item)
short_supply_detail['主站名称'] = short_supply_detail['城市仓'].map(mapping)
check_list = list(set(short_supply_detail['责任方']))
mapping = {
	'调拨差异': '调拨差异', '仓配': '仓储', '配送': '配送', '其他': '其他', '仓储': '仓储', '配货': '仓储', '仓库': '仓储',
	'品控': '品控', '超卖': '超卖', '采购': '采购', '运营': '运营', '不配送': '仓储', '到货不足': '供应商', '供应商': '供应商',
	'溧阳仓': '仓储', '杭州仓': '仓储', '温州仓': '仓储', '南京仓': '仓储', '加工': '仓储', '数据': '仓储', '出库': '仓储',
	'今天带回': '配送', '溧阳仓库': '仓储', '溧阳调拨': '仓储', '退款': '运营', '司机': '配送', '供应商带走': '供应商', '加工中心': '加工中心',
	'市场': '市场', 'TC': 'TC', 'FC': '仓储', '售后': '售后', 'SC': '仓储', '未标记': '未标记', '客服': '客服'}
for item in check_list:
	if item not in mapping:
		print('责任方', item)
# 字段扩展_汇总
short_supply_detail['责任方'] = short_supply_detail['责任方'].map(mapping)
short_supply_detail = short_supply_detail.drop('城市仓', axis=1)
short_supply_detail_all = short_supply_detail.groupby(['日期', '商品名称', '缺货原因', '责任方', '采购'], as_index=False)['缺货件数'].sum()
short_supply_detail_all['主站名称'] = '华东'
short_supply = short_supply_detail.groupby(['主站名称', '日期', '缺货原因', '责任方', '采购'], as_index=False)['缺货件数'].sum()
total_shipments_only = total_shipments[['主站名称', '日期']].drop_duplicates()
short_supply_only = short_supply[['缺货原因', '责任方', '采购']].drop_duplicates()
total_shipments_part = np.tile(total_shipments_only.values, (short_supply_only.shape[0], 1))
short_supply_part = np.repeat(short_supply_only.values, total_shipments_only.shape[0], axis=0)
part = np.concatenate([total_shipments_part, short_supply_part], axis=1)
part_df = pd.DataFrame(part, columns=['主站名称', '日期', '缺货原因', '责任方', '采购'])
part_df['缺货件数'] = 0
short_supply = pd.concat([short_supply, part_df], axis=0, ignore_index=True)
short_supply = short_supply.groupby(['主站名称', '日期', '缺货原因', '责任方', '采购'], as_index=False)['缺货件数'].sum()

print(short_supply['缺货件数'].sum())
print(short_supply_detail['缺货件数'].sum())
short_supply_detail['缺货件数'] = short_supply_detail['缺货件数'].astype(int)

short_supply_detail_reason = short_supply_detail.groupby(['主站名称', '日期', '商品名称', '责任方'], as_index=False)['缺货件数'].sum()
short_supply_detail_reason = short_supply_detail_reason.sort_values(['主站名称', '日期', '商品名称', '缺货件数'], ascending=False)
short_supply_detail_reason = short_supply_detail_reason.drop_duplicates(['主站名称', '日期', '商品名称'])
short_supply_detail_reason['主要责任方'] = short_supply_detail_reason['责任方'] + '/' + short_supply_detail_reason['缺货件数'].astype(str)
short_supply_detail_reason = short_supply_detail_reason[['主站名称', '日期', '商品名称', '主要责任方']]
short_supply_detail = short_supply_detail.merge(short_supply_detail_reason, on=['主站名称', '日期', '商品名称'], how='left')

# short_supply_detail_all_reason = short_supply_detail_all.groupby(['主站名称', '日期', '商品名称', '责任方'], as_index=False)['缺货件数'].sum()
# short_supply_detail_all_reason = short_supply_detail_all_reason.sort_values(['主站名称', '日期', '商品名称', '缺货件数'], ascending=False)
# short_supply_detail_all_reason = short_supply_detail_all_reason.drop_duplicates(['主站名称', '日期', '商品名称'])
# short_supply_detail_all_reason['主要责任方'] = short_supply_detail_all_reason['责任方'] + '/' + short_supply_detail_all_reason['缺货件数'].astype(str)
# short_supply_detail_all_reason = short_supply_detail_all_reason[['主站名称', '日期', '商品名称', '主要责任方']]
# short_supply_detail_all = short_supply_detail_all.merge(short_supply_detail_all_reason, on=['主站名称', '日期', '商品名称'], how='left')

short_supply.to_csv(r'D:\文档\工作\十荟团\temp\temp_short_supply.csv', index=False, encoding='utf_8_sig')
short_supply_detail.to_csv(r'D:\文档\工作\十荟团\temp\temp_short_supply_detail.csv', index=False, encoding='utf_8_sig')
# short_supply_detail_all.to_csv(r'D:\文档\工作\十荟团\temp\temp_short_supply_detail_all.csv', index=False, encoding='utf_8_sig')

sys.exit(0)