import pandas as pd
import re
import warnings
from tqdm import tqdm
import math
from multiprocessing import Pool, cpu_count
from difflib import SequenceMatcher
import numpy as np
import sys
import datetime

pd.options.display.max_columns = None
pd.options.display.max_rows = None
warnings.filterwarnings('ignore')


def list_add(x):
	r = []
	for l in x:
		for item in l:
			r.append(item)
	return r


def abnormal_tuan(df, result_type):
	df['订单日期'] = pd.to_datetime(df['订单日期'])
	df['date'] = df['订单日期'].dt.date
	df['month'] = df['订单日期'].dt.month
	df['year'] = df['订单日期'].dt.year
	df['日期'] = df['year'].astype(str) + '-' + df['month'].astype(str)
	proporation_threshold = 0.8

	# 团内多个用户对同一商品刷单,增加一小时的时间窗口
	temp_product = df.groupby(['团长id', 'date', '购买人用户id'], as_index=False)['商品id'].agg({'product_nunique': 'nunique'})
	temp_user = df.copy()
	temp_user = temp_user.merge(temp_product, on=['团长id', 'date', '购买人用户id'], how='left')
	temp_user = temp_user[temp_user['product_nunique'] == 1]
	temp_user = temp_user.groupby(['团长id', 'date', '商品id'], as_index=False)['购买人用户id'].agg({'user_nunique': 'nunique'})
	df = df.merge(temp_product, on=['团长id', 'date', '购买人用户id'], how='left')
	df = df.merge(temp_user, on=['团长id', 'date', '商品id'], how='left')
	df['abnormal'] = df.apply(lambda x: 1 if (x['product_nunique'] == 1) and (x['user_nunique'] >= 5) else 0, axis=1)
	temp = df[df['abnormal'] == 1]
	temp = temp.groupby(['团长id', 'date', '商品id'], as_index=False)['下单时间2'].agg({'max_time': 'max', 'min_time': 'min'})
	temp['time_gap'] = temp.apply(lambda x: (x['max_time'] - x['min_time']).total_seconds()/3600, axis=1)
	temp = temp.drop(['max_time', 'min_time'], axis=1)
	df = df.merge(temp, on=['团长id', 'date', '商品id'], how='left')
	df['abnormal'] = df.apply(lambda x: 1 if (x['abnormal'] == 1) and (x['time_gap'] <= 1) else 0, axis=1)
	# 一个用户对单一商品数量超过阈值,20210907阈值20-25
	temp_product = df.groupby(['date', '购买人用户id', '商品id'], as_index=False)['数量'].agg({'总数量': 'sum'})
	df = df.merge(temp_product, on=['date', '购买人用户id', '商品id'], how='left')
	df['abnormal'] = df.apply(lambda x: 1 if x['总数量'] >= 25 else x['abnormal'], axis=1)
	# 一个用户对单一商品销售额超过阈值
	temp_product = df.groupby(['date', '购买人用户id', '商品id'], as_index=False)['销售额'].agg({'总销售额': 'sum'})
	df = df.merge(temp_product, on=['date', '购买人用户id', '商品id'], how='left')
	df['abnormal'] = df.apply(lambda x: 1 if (x['总销售额'] >= 500) and (x['总数量'] >= 2) else x['abnormal'], axis=1)
	# 一个用户对单一商品订单数超过阈值
	temp_product = df.groupby(['date', '购买人用户id', '商品id'], as_index=False)['订单id'].agg({'订单数': 'nunique'})
	df = df.merge(temp_product, on=['date', '购买人用户id', '商品id'], how='left')
	df['abnormal'] = df.apply(lambda x: 1 if x['订单数'] >= 5 else x['abnormal'], axis=1)
	temp = df[df['abnormal'] == 1]

	if result_type == 'month':
		total_sales = df.groupby(['团长id', '日期'], as_index=False).agg({'数量': 'sum', '销售额': 'sum', '购买人用户id': 'nunique'})
		total_sales = total_sales.rename(columns={'购买人用户id': 'user_nunique_month'})
		temp = temp.groupby(['团长id', '日期'], as_index=False).agg({'数量': 'sum', '销售额': 'sum', '购买人用户id': 'nunique'})
		temp = temp.rename(columns={'数量': '数量_abnormal', '销售额': '销售额_abnormal', '购买人用户id': 'user_nunique_month_abnormal'})
		total_sales = total_sales.merge(temp, on=['团长id', '日期'], how='left')
		for col in ['数量_abnormal', '销售额_abnormal', 'user_nunique_month_abnormal']:
			total_sales[col] = total_sales[col].fillna(0)
			total_sales[f'{col}_ration'] = total_sales[col] / total_sales[re.sub('_abnormal', '', col)]
		result = total_sales[(total_sales['数量'] >= 10) & (total_sales['销售额'] >= 500) & (total_sales['user_nunique_month'] >= 20)]
		result['result_1'] = result.apply(lambda x: [f'{x["日期"]}月异常数量占比{round(x["数量_abnormal_ration"] * 100, 1)}%'] if x["数量_abnormal_ration"]>proporation_threshold else list(), axis=1)
		result['result_2'] = result.apply(lambda x: [f'{x["日期"]}月异常销售额占比{round(x["销售额_abnormal_ration"] * 100, 1)}%'] if x["销售额_abnormal_ration"]>proporation_threshold else list(), axis=1)
		# result['result_3'] = result.apply(lambda x: [f'{x["日期"]}月异常用户数占比{round(x["user_nunique_month_abnormal_ration"] * 100, 1)}%'] if x["user_nunique_month_abnormal_ration"]>proporation_threshold else list(), axis=1)
		result['caution'] = result['result_1'] + result['result_2']
		# result.to_csv('C:/Users/aesdhj/Desktop/result_origin_month.csv', index=False, encoding='utf_8_sig')
		# result = result[result['caution'].apply(len).gt(0)]
		result = result.groupby('团长id')['caution'].apply(lambda x: list_add(list(x))).reset_index()
	elif result_type == 'all':
		total_sales = df.groupby(['团长id'], as_index=False).agg({'数量': 'sum', '销售额': 'sum', '购买人用户id': 'nunique'})
		total_sales = total_sales.rename(columns={'购买人用户id': 'user_nunique'})
		temp = temp.groupby(['团长id'], as_index=False).agg({'数量': 'sum', '销售额': 'sum', '购买人用户id': 'nunique'})
		temp = temp.rename(columns={'数量': '数量_abnormal', '销售额': '销售额_abnormal', '购买人用户id': 'user_nunique_abnormal'})
		total_sales = total_sales.merge(temp, on=['团长id'], how='left')
		for col in ['数量_abnormal', '销售额_abnormal', 'user_nunique_abnormal']:
			total_sales[col] = total_sales[col].fillna(0)
			total_sales[f'{col}_ration'] = total_sales[col] / total_sales[re.sub('_abnormal', '', col)]
		result = total_sales[(total_sales['数量'] >= 10) & (total_sales['销售额'] >= 500) & (total_sales['user_nunique'] >= 20)]
		result['result_1'] = result.apply(lambda x: [f'异常数量占比{round(x["数量_abnormal_ration"] * 100, 1)}%'] if x["数量_abnormal_ration"]>proporation_threshold else list(), axis=1)
		result['result_2'] = result.apply(lambda x: [f'异常销售额占比{round(x["销售额_abnormal_ration"] * 100, 1)}%'] if x["销售额_abnormal_ration"]>proporation_threshold else list(), axis=1)
		# result['result_3'] = result.apply(lambda x: [f'异常用户数占比{round(x["user_nunique_abnormal_ration"] * 100, 1)}%'] if x["user_nunique_abnormal_ration"]>proporation_threshold else list(), axis=1)
		result['caution'] = result['result_1'] + result['result_2']
		# result.to_csv('C:/Users/aesdhj/Desktop/result_origin.csv', index=False, encoding='utf_8_sig')
		# result = result[result['caution'].apply(len).gt(0)]
		result = result.groupby('团长id')['caution'].apply(lambda x: list_add(list(x))).reset_index()
	return result


# 团内长期来看消费行为相似相关联用户组
# 20210920增加跨团数据和微信昵称相似性
def tuan_in_user_action_abnormal(df):
	df['订单日期'] = pd.to_datetime(df['订单日期'])
	df['date'] = df['订单日期'].dt.date
	df['month'] = df['订单日期'].dt.month
	df['year'] = df['订单日期'].dt.year
	df['product'] = df['date'].astype(str) + '_' + df['商品id'].astype(str)
	proporation_threshold = 0.8
	sim_threshold = 0.5

	tuan_list = list(set(df['团长id']))
	user_group_list = []
	for tuan_id in tqdm(tuan_list):
		temp_origin = df[df['团长id'] == tuan_id]
		user_origin = list(set(temp_origin['购买人用户id']))
		temp = df[df['购买人用户id'].isin(user_origin)]

		# 筛选条件
		temp_sum = temp.groupby(['购买人用户id'], as_index=False)['数量', '销售额'].sum()
		temp_sum = temp_sum.rename(columns={'数量': '总数量', '销售额': '总销售额'})
		temp = temp.merge(temp_sum, on=['购买人用户id'], how='left')
		temp_x = temp[~((temp['总数量'] <= 25) & (temp['总销售额'] <= 500))]

		user_list = list(set(temp_x['购买人用户id']))
		user_product = temp_x.groupby('购买人用户id')['product'].agg(list).reset_index()
		user_product_dict = dict(zip(user_product['购买人用户id'], user_product['product']))
		# 20210917更新加入了用户微信昵称相似性
		temp_name = temp[['购买人用户id', '购买人微信昵称']].drop_duplicates('购买人用户id')
		temp_name = temp_name[~pd.isna(temp_name['购买人微信昵称'])]
		user_name_dict = dict(zip(temp_name['购买人用户id'], temp_name['购买人微信昵称']))
		user_action_sim = {}
		user_list_len = len(user_list)
		for i in range(0, user_list_len):
			user_action_sim.setdefault(user_list[i], {})
			user_p = set(user_product_dict[user_list[i]])
			for j in range(i+1, user_list_len):
				user_related_p = set(user_product_dict[user_list[j]])
				sim_action = len(user_p & user_related_p) / math.sqrt(len(user_p) * len(user_related_p))
				if (user_list[i] in user_name_dict) and (user_list[j] in user_name_dict):
					sim_name = max(
						SequenceMatcher(None, user_name_dict[user_list[i]], user_name_dict[user_list[j]]).ratio(),
						SequenceMatcher(None, user_name_dict[user_list[j]], user_name_dict[user_list[i]]).ratio())
					if sim_name >= 0.6:
						sim = max(sim_action, sim_name)
					else:
						sim = sim_action
				else:
					sim = sim_action
				if sim < sim_threshold:
					continue
				user_action_sim[user_list[i]][user_list[j]] = sim

		user_action_sim_new = {}
		for k, v in user_action_sim.items():
			if len(v) > 0:
				user_action_sim_new[k] = v

		user_group = temp_origin.groupby(['团长id', '购买人用户id'], as_index=False)['数量', '销售额'].sum()
		user_group['user_nunique'] = 1
		user_group['group'] = -1
		group_index = 1
		for id, sim_group in user_action_sim_new.items():
			for id_related, sim in sim_group.items():
				if sim >= sim_threshold:
					if user_group[user_group['购买人用户id'] == id]['group'].values[0] == -1 and user_group[user_group['购买人用户id'] == id_related]['group'].values[0] == -1:
						user_group.loc[user_group['购买人用户id'] == id, 'group'] = group_index
						user_group.loc[user_group['购买人用户id'] == id_related, 'group'] = group_index
						group_index += 1
					elif user_group[user_group['购买人用户id'] == id]['group'].values[0] == -1 and user_group[user_group['购买人用户id'] == id_related]['group'].values[0] != -1:
						user_group.loc[user_group['购买人用户id'] == id, 'group'] = user_group[user_group['购买人用户id'] == id_related]['group'].values[0]
					elif user_group[user_group['购买人用户id'] == id]['group'].values[0] != -1 and user_group[user_group['购买人用户id'] == id_related]['group'].values[0] == -1:
						user_group.loc[user_group['购买人用户id'] == id_related, 'group'] = user_group[user_group['购买人用户id'] == id]['group'].values[0]
					elif user_group[user_group['购买人用户id'] == id]['group'].values[0] != -1 and user_group[user_group['购买人用户id'] == id_related]['group'].values[0] != -1:
						user_group.loc[user_group['group'] == user_group[user_group['购买人用户id'] == id]['group'].values[0], 'group'] = user_group[user_group['购买人用户id'] == id_related]['group'].values[0]
		user_group = user_group.groupby(['团长id', 'group'], as_index=False)['数量', '销售额', 'user_nunique'].sum()
		user_group_list.append(user_group)

	user_group_all = pd.concat(user_group_list, axis=0, ignore_index=True)
	user_group_all.to_csv(f'C:/Users/aesdhj/Desktop/user_group_all.csv', index=False, encoding='utf_8_sig')
	total_sales = user_group_all.groupby('团长id', as_index=False)['数量', '销售额', 'user_nunique'].sum()
	user_group_all = user_group_all.rename(columns={'数量': '数量_abnormal', '销售额': '销售额_abnormal', 'user_nunique': 'user_nunique_abnormal'})
	user_group_all = user_group_all.merge(total_sales, on='团长id', how='left')
	for col in ['数量_abnormal', '销售额_abnormal', 'user_nunique_abnormal']:
		user_group_all[col] = user_group_all[col].fillna(0)
		user_group_all[f'{col}_ration'] = user_group_all[col] / user_group_all[re.sub('_abnormal', '', col)]
	user_group_all['result_1'] = user_group_all.apply(lambda x: [f'异常数量占比{round(x["数量_abnormal_ration"] * 100, 1)}%'] if (x["数量_abnormal_ration"]>proporation_threshold) and (x['group'] != -1) else list(), axis=1)
	user_group_all['result_2'] = user_group_all.apply(lambda x: [f'异常销售额占比{round(x["销售额_abnormal_ration"] * 100, 1)}%'] if (x["销售额_abnormal_ration"]>proporation_threshold) and (x['group'] != -1) else list(), axis=1)
	# user_group_all['result_3'] = user_group_all.apply(lambda x: [f'异常用户数占比{round(x["user_nunique_abnormal_ration"] * 100, 1)}%'] if (x["user_nunique_abnormal_ration"]>proporation_threshold) and (x['group'] != -1) else list(), axis=1)
	user_group_all['caution'] = user_group_all['result_1'] + user_group_all['result_2']
	# user_group_all.to_csv('C:/Users/aesdhj/Desktop/result_action_part_{}.csv'.format(hash(str(user_product_dict))), index=False, encoding='utf_8_sig')
	result = user_group_all.groupby('团长id')['caution'].apply(lambda x: list_add(list(x))).reset_index()
	return result


if __name__ == '__main__':

	df = pd.read_csv(r'C:\Users\aesdhj\Desktop\df_jiangsu.csv')
	df['团长id'] = df['团长id'].astype(int)
	df['购买人用户id'] = df['购买人用户id'].astype(int)
	df['订单id'] = df['订单id'].astype(str)
	df['下单时间2'] = pd.to_datetime(df['下单时间2'])
	temp = df[['购买人用户id', '城市圈']]
	temp = temp[~(pd.isna(temp['城市圈']) | (temp['城市圈'] == '包邮到家'))]
	temp = temp.drop_duplicates('购买人用户id')
	temp_dict = dict(zip(temp['购买人用户id'], temp['城市圈']))
	df['城市圈'] = df.apply(lambda x: temp_dict[x['购买人用户id']] if (x['购买人用户id'] in temp_dict) and (pd.isna(x['城市圈']) or (x['城市圈'] == '包邮到家')) else x['城市圈'], axis=1)
	temp = df[['购买人用户id', '购买人微信昵称']]
	temp = temp[~pd.isna(temp['购买人微信昵称'])]
	temp = temp.drop_duplicates('购买人用户id')
	temp_dict = dict(zip(temp['购买人用户id'], temp['购买人微信昵称']))
	df['购买人微信昵称'] = df.apply(lambda x: temp_dict[x['购买人用户id']] if (x['购买人用户id'] in temp_dict) and pd.isna(x['购买人微信昵称']) else x['购买人微信昵称'], axis=1)


	# 团内不同用户对在不同日期刷单同一商品，透视表对角形式&随机刷单
	result_type = 'all'
	result = abnormal_tuan(df, result_type)
	result.to_csv(f'C:/Users/aesdhj/Desktop/result_{result_type}.csv', index=False, encoding='utf_8_sig')
	print(f'result_type_{result_type} done')


	# 团内同一批用户对在不同日期刷单不同商品，透视表横排形式
	if len(set(df['团长id'])) < cpu_count() - 2:
		split_num = 1
	else:
		split_num = cpu_count() - 2
	tuan_list = list(set(df['团长id']))
	size = int(math.ceil(len(tuan_list) / split_num))
	df_list = []
	for i in range(split_num):
		start = i * size
		end = (i + 1) * size if (i + 1) * size < len(tuan_list) else len(tuan_list)
		tuan_list_part = tuan_list[start: end]
		temp = df[df['团长id'].isin(tuan_list_part)]
		df_list.append(temp)
	p = Pool(split_num)
	rlist = p.map(tuan_in_user_action_abnormal, df_list)
	p.close()
	p.join()
	result_action = pd.concat(rlist, axis=0, ignore_index=True)
	result_action.to_csv('C:/Users/aesdhj/Desktop/result_action.csv', index=False, encoding='utf_8_sig')
	print('result_action done')





