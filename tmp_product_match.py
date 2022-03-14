from difflib import SequenceMatcher
import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None
import numpy as np
import sys
import re

def clear_product(product):
	product_clear = re.sub('【.*?】', '', product)
	product_clear = re.sub('[\s+]', '', product_clear)
	product_clear = product_clear.replace('颜色随机', '')
	product_clear = product_clear.replace('款式随机', '')
	return product_clear

def product_match(product, all_product, product_upc_list, product_list):

	match_list = []
	for item in product_upc_list:
		item = clear_product(item)
		match_list.append(SequenceMatcher(None, product, item).ratio())
	product_upc_match = product_list[np.argmax(match_list)]
	product_id_upc_match = all_product[all_product['商品'] == product_upc_match]['ID'].values[0]
	product_id_upc_ration = match_list[int(np.argmax(match_list))]
	product_upc_match_clear = clear_product(product_upc_match)
	if product in product_upc_match_clear:
		product_id_upc_ration = 1

	match_list = []
	for item in product_list:
		item = clear_product(item)
		match_list.append(SequenceMatcher(None, product, item).ratio())
	product_match = product_list[np.argmax(match_list)]
	product_id_match = all_product[all_product['商品'] == product_match]['ID'].values[0]
	product_id_ration = match_list[int(np.argmax(match_list))]
	product_match_clear = clear_product(product_match)
	if product in product_match_clear:
		product_id_ration = 1
	return [product_id_upc_match, product_upc_match, product_id_upc_ration, product_id_match, product_match, product_id_ration]


def product_match_new(product, all_product, product_list):
	match_list = []
	for item in product_list:
		item = clear_product(item)
		match_list.append(SequenceMatcher(None, product, item).ratio())
	product_match = product_list[np.argmax(match_list)]
	product_id_match = all_product[all_product['商品'] == product_match]['ID'].values[0]
	product_id_ration = match_list[int(np.argmax(match_list))]
	product_match_clear = clear_product(product_match)
	if product in product_match_clear:
		product_id_ration = 1

	match_list[np.argmax(match_list)] = -1
	product_match_lower = product_list[np.argmax(match_list)]
	product_id_match_lower = all_product[all_product['商品'] == product_match_lower]['ID'].values[0]
	product_id_ration_lower = match_list[int(np.argmax(match_list))]
	product_match_lower_clear = clear_product(product_match_lower)
	if product in product_match_lower_clear:
		product_id_ration_lower = 1

	return [product_id_match_lower, product_match_lower, product_id_ration_lower, product_id_match, product_match, product_id_ration]


touxiandan_jiangsu = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_jiangsu.xlsx', sheet_name=None)
tc_different_jiangsu = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异jiangsu.xlsx')
# touxiandan_wenzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_wenzhou.xlsx', sheet_name=None)
# tc_different_wenzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异wenzhou.xlsx')
# touxiandan_hanzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_hanzhou.xlsx', sheet_name=None)
# tc_different_hanzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异hanzhou.xlsx')
# touxiandan_hefei = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_hefei.xlsx', sheet_name=None)
# tc_different_hefei = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异hefei.xlsx')
# touxiandan_xuzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_xuzhou.xlsx', sheet_name=None)
# tc_different_xuzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异xuzhou.xlsx')

for items in [
	(touxiandan_jiangsu, tc_different_jiangsu, 'jiangsu'),
	# (touxiandan_wenzhou, tc_different_wenzhou, 'wenzhou'),
	# (touxiandan_hefei, tc_different_hefei, 'hefei'),
	# (touxiandan_xuzhou, tc_different_xuzhou, 'xuzhou'),
	# (touxiandan_hanzhou, tc_different_hanzhou, 'hanzhou'),
]:
	touxiandan = items[0]
	tc_different = items[1]
	all_product = touxiandan['1-汇总']
	file_name = items[2]

	# 未提报数据TC
	tc_lost = []
	tc_all = [c for c in all_product.columns if 'tc' in c] + [c for c in all_product.columns if 'TC' in c]
	for item in tc_all:
		if item not in list(set(tc_different['服务站'])) + ['溧阳-宜兴TC', '溧阳-常州TC', '溧阳-溧阳TC', '溧阳-郎溪TC', '新北LTC', '杭州-嘉兴直配', '溧阳LTC', '虚拟TC', '上海嘉定LTC', '杭州市区直配']:
			tc_lost.append(item)
	if len(tc_lost) > 0:
		print(tc_lost)
		# sys.exit(0)

	cols = list(all_product.columns)
	cols[1] = '商品'
	all_product.columns = cols
	all_product = all_product[['商品', 'ID', '条形码', '分类']]
	all_product = all_product[~pd.isna(all_product['ID'])]
	all_product['ID'] = all_product['ID'].astype(int)
	all_product['ID'] = all_product['ID'].astype(str)
	all_product['条形码'] = all_product['条形码'].astype(str)
	all_product = all_product.drop_duplicates()
	all_product['商品_条形码'] = all_product['商品'] + all_product['条形码']
	product_list = list(all_product['商品'])
	# product_upc_list = list(all_product['商品_条形码'])

	tc_different['sku或规格Id'] = tc_different['sku或规格Id'].astype(int)
	tc_different['sku或规格Id'] = tc_different['sku或规格Id'].astype(str)
	# tc_different['match_list'] = tc_different.apply(lambda x: product_match(x['商品名称'], all_product, product_upc_list, product_list) if x['sku或规格Id'] == '0' else ['-', '-', 0, '-', '-', 0], axis=1)
	tc_different['match_list'] = tc_different.apply(lambda x: product_match_new(x['商品名称'], all_product, product_list) if x['sku或规格Id'] == '0' else ['-', '-', 0, '-', '-', 0], axis=1)
	tc_different['id'] = tc_different['match_list'].apply(lambda x: x[3])
	tc_different['product'] = tc_different['match_list'].apply(lambda x: x[4])
	tc_different['product_ration'] = tc_different['match_list'].apply(lambda x: x[5])
	tc_different['id_lower'] = tc_different['match_list'].apply(lambda x: x[0])
	tc_different['product_lower'] = tc_different['match_list'].apply(lambda x: x[1])
	tc_different['product_ration_lower'] = tc_different['match_list'].apply(lambda x: x[2])
	tc_different['sku或规格Id'] = tc_different.apply(lambda x: x['id'] if x['product_ration'] >= 0.6 else x['sku或规格Id'], axis=1)
	tc_different['商品名称'] = tc_different.apply(lambda x: x['product'] if x['product_ration'] >= 0.6 else x['商品名称'], axis=1)
	tc_different['sku或规格Id'] = tc_different.apply(lambda x: x['id_lower'] if (x['product_ration'] < 0.6) and (x['product_ration_lower'] >= 0.6) else x['sku或规格Id'], axis=1)
	tc_different['商品名称'] = tc_different.apply(lambda x: x['product_lower'] if (x['product_ration'] < 0.6) and (x['product_ration_lower'] >= 0.6) else x['商品名称'], axis=1)
	tc_different = tc_different.drop('match_list', axis=1)
	tc_different.to_excel('C:/Users/aesdhj/Desktop/服务站差异_adjusted_{}.xlsx'.format(file_name), sheet_name='服务站差异0', index=False)

sys.exit(0)