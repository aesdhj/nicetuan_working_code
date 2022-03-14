import pandas as pd
pd.options.display.max_columns = None
pd.options.display.max_rows = None
import warnings
warnings.filterwarnings('ignore')
from datetime import datetime
import sys
import re
from datetime import datetime, timedelta, date
mapping_ltc = {
	'武进直配': '武进LTC', '无锡直配': '无锡LTC', '上海嘉定直配': '上海嘉定LTC', '安徽直配': '安徽LTC', '溧阳直配1': '溧阳LTC'
}

short_supply_detail_tc_jiangsu = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_jiangsu.xlsx', sheet_name=None)
short_supply_detail_imported_jiangsu = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异_adjusted_jiangsu.xlsx')
arrival_time_jiangsu = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站支线jiangsu.xlsx')
# short_supply_detail_tc_wenzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_wenzhou.xlsx', sheet_name=None)
# short_supply_detail_imported_wenzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异_adjusted_wenzhou.xlsx')
# arrival_time_wenzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站支线wenzhou.xlsx')
# short_supply_detail_tc_hanzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_hanzhou.xlsx', sheet_name=None)
# short_supply_detail_imported_hanzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异_adjusted_hanzhou.xlsx')
# arrival_time_hanzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站支线hanzhou.xlsx')
# short_supply_detail_tc_hefei = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_hefei.xlsx', sheet_name=None)
# short_supply_detail_imported_hefei = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异_adjusted_hefei.xlsx')
# arrival_time_hefei = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站支线hefei.xlsx')
# short_supply_detail_tc_xuzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\touxiandan_xuzhou.xlsx', sheet_name=None)
# short_supply_detail_imported_xuzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站差异_adjusted_xuzhou.xlsx')
# arrival_time_xuzhou = pd.read_excel(r'C:\Users\aesdhj\Desktop\服务站支线xuzhou.xlsx')

for items in [
	(short_supply_detail_tc_jiangsu, short_supply_detail_imported_jiangsu, 'jiangsu', arrival_time_jiangsu),
	# (short_supply_detail_tc_wenzhou, short_supply_detail_imported_wenzhou, 'wenzhou', arrival_time_wenzhou),
	# (short_supply_detail_tc_hefei, short_supply_detail_imported_hefei, 'hefei', arrival_time_hefei),
	# (short_supply_detail_tc_xuzhou, short_supply_detail_imported_xuzhou, 'xuzhou', arrival_time_xuzhou),
	# (short_supply_detail_tc_hanzhou, short_supply_detail_imported_hanzhou, 'hanzhou', arrival_time_hanzhou),
]:
	short_supply_detail_tc = items[0]
	short_supply_detail_imported = items[1]
	file_name = items[2]
	arrival_time = items[3]
	short_supply_detail_imported['缺货数量'] = short_supply_detail_imported.apply(lambda x: x['多货数量'] if (x['备注'] == '破损') and (x['多货数量'] > 0) else x['缺货数量'], axis=1)
	short_supply_detail_imported['多货数量'] = short_supply_detail_imported.apply(lambda x: 0 if (x['备注'] == '破损') and (x['多货数量'] > 0) else x['多货数量'], axis=1)

	total_supply = short_supply_detail_tc['1-汇总']
	cols = list(total_supply.columns)
	cols[1] = '商品'
	total_supply.columns = cols
	total_supply['商品'] = total_supply.apply(lambda x: x['商品'].replace('\n', ''), axis=1)
	total_supply = total_supply[~pd.isna(total_supply['序号'])]
	total_supply['ID'] = total_supply['ID'].astype(int)
	total_supply['ID'] = total_supply['ID'].astype(str)
	total_supply = total_supply.drop(['序号', '条形码', '货位', '总计'], axis=1)
	trans_cols = [c for c in total_supply.columns if 'TC' in c] + [c for c in total_supply.columns if 'tc' in c]
	for c in trans_cols:
		total_supply[c] = total_supply[c].fillna(0)
		total_supply[c] = total_supply[c].astype(int)
	total_supply = total_supply.set_index(['商品', 'ID', '分类']).stack().reset_index()
	total_supply = total_supply.rename(columns={'level_3': 'TC', 0: 'TC应出库总件数'})
	total_supply = total_supply.rename(columns={'ID': '规格ID'})
	total_supply['缺货数量'] = 0
	total_supply['缺货数量_TC'] = 0
	total_supply['缺货数量_FC'] = 0
	total_supply['缺货数量'] = 0
	total_supply['多货数量'] = 0
	id_name = {}
	id_category = {}         
	for item in total_supply[['规格ID', '商品', '分类']].values:
		id_name[item[0]] = item[1]
		id_category[item[0]] = item[2]
	# 20210525更新，TC名称变化
	total_supply['TC'] = total_supply.apply(lambda x: re.sub('\d+', '', x['TC']), axis=1)

	# arrival_time = arrival_time[arrival_time['类型'] == '到站']
	# arrival_time = arrival_time[~pd.isna(arrival_time['打卡时间'])]
	# arrival_time['打卡时间'] = pd.to_datetime(arrival_time['打卡时间'])
	# arrival_time_latest = arrival_time.groupby('服务站', as_index=False)['打卡时间'].max()
	# arrival_time_latest['服务站'] = arrival_time_latest.apply(lambda x: re.sub('\d+', '', x['服务站']), axis=1)
	# arrival_time_latest = arrival_time_latest.rename(columns={'服务站': 'TC'})
	# arrival_time_latest['latest_3hour_later'] = arrival_time_latest['打卡时间'] + timedelta(hours=3)
	# arrival_time_latest = arrival_time_latest[['TC', 'latest_3hour_later']]
	# 2021-08-18
	arrival_time['tc_count'] = arrival_time.apply(lambda x: len(x['TC'].split(',')), axis=1)
	temp = arrival_time[arrival_time['tc_count'] >= 2]
	arrival_time = arrival_time[arrival_time['tc_count'] < 2]
	cols = temp.columns
	data_list = []
	for item in temp.values:
		tc_count = item[-1]
		for i in range(tc_count):
			item_copy = item.copy()
			item_copy[4] = item[4].split(',')[i]
			data_list.append(item_copy)
	temp = pd.DataFrame(data_list, columns=cols)
	arrival_time = pd.concat([arrival_time, temp], axis=0, ignore_index=True)
	# 2021-08-06
	arrival_time = arrival_time[arrival_time['打卡类型'] == '到站']
	arrival_time = arrival_time[~pd.isna(arrival_time['到站打卡时间'])]
	arrival_time['到站打卡时间'] = pd.to_datetime(arrival_time['到站打卡时间'])
	arrival_time_latest = arrival_time.groupby('TC', as_index=False)['到站打卡时间'].max()
	arrival_time_latest['TC'] = arrival_time_latest.apply(lambda x: re.sub('\d+', '', x['TC']), axis=1)
	arrival_time_latest['latest_3hour_later'] = arrival_time_latest['到站打卡时间'] + timedelta(hours=3)
	arrival_time_latest = arrival_time_latest[['TC', 'latest_3hour_later']]


	short_supply_detail_imported = short_supply_detail_imported.rename(columns={'服务站': 'TC', 'sku或规格Id': '规格ID', '商品名称': '商品'})
	# 20210525更新，TC名称变化
	short_supply_detail_imported['TC'] = short_supply_detail_imported.apply(lambda x: re.sub('\d+', '', x['TC']), axis=1)
	short_supply_detail_imported['规格ID'] = short_supply_detail_imported['规格ID'].astype(str)
	short_supply_detail_imported = short_supply_detail_imported.merge(arrival_time_latest, on='TC', how='left')
	short_supply_detail_imported['提报时间'] = pd.to_datetime(short_supply_detail_imported['提报时间'])
	short_supply_detail_imported['缺货类型'] = short_supply_detail_imported.apply(lambda x:
									'TC' if pd.isna(x['latest_3hour_later']) else
									('TC' if pd.isna(x['提报时间']) and ~pd.isna(x['latest_3hour_later']) else
									('FC' if x['提报时间'] <= x['latest_3hour_later'] else 'TC'))
									, axis=1)
	short_supply_detail_imported['缺货数量_FC'] = short_supply_detail_imported.apply(lambda x: x['缺货数量'] if x['缺货类型'] == 'FC' else 0, axis=1)
	short_supply_detail_imported['缺货数量_TC'] = short_supply_detail_imported.apply(lambda x: x['缺货数量'] if x['缺货类型'] == 'TC' else 0, axis=1)

	short_supply_detail_imported_damaged = short_supply_detail_imported[short_supply_detail_imported['备注'] == '破损']
	short_supply_detail_imported = short_supply_detail_imported[['商品', '规格ID', '分类', 'TC', '缺货数量', '多货数量', '缺货数量_FC', '缺货数量_TC']]
	short_supply_detail_imported['TC应出库总件数'] = 0
	short_supply_detail_imported['商品'] = short_supply_detail_imported.apply(lambda x: id_name[x['规格ID']] if x['规格ID'] in id_name else x['商品'], axis=1)
	short_supply_detail_imported['分类'] = short_supply_detail_imported.apply(lambda x: id_category[x['规格ID']] if x['规格ID'] in id_category else x['分类'], axis=1)
	short_supply_detail_imported['分类'] = short_supply_detail_imported['分类'].fillna('add')

	tc_short_supply_detail = pd.concat([total_supply, short_supply_detail_imported], axis=0, ignore_index=True)
	tc_short_supply_detail = tc_short_supply_detail.groupby(['TC', '规格ID', '商品', '分类'], as_index=False)['TC应出库总件数', '多货数量', '缺货数量', '缺货数量_FC', '缺货数量_TC'].sum()

	short_supply_detail_imported_damaged = short_supply_detail_imported_damaged.groupby(['TC', '规格ID'], as_index=False)['缺货数量'].sum()
	short_supply_detail_imported_damaged = short_supply_detail_imported_damaged.rename(columns={'缺货数量': '破损'})
	tc_short_supply_detail = tc_short_supply_detail.merge(short_supply_detail_imported_damaged, on=['TC', '规格ID'], how='left')
	tc_short_supply_detail['破损'] = tc_short_supply_detail['破损'].fillna(0)
	tc_short_supply_detail['TC'] = tc_short_supply_detail.apply(lambda x: mapping_ltc[x['TC']] if x['TC'] in mapping_ltc else x['TC'], axis=1)
	tc_short_supply_detail.to_csv('D:/文档/工作/十荟团/temp/华东TC缺货改期_新/{}_{}.csv'.format(datetime.now().strftime('%Y-%m-%d'), file_name), index=False, encoding='utf_8_sig')

sys.exit(0)
