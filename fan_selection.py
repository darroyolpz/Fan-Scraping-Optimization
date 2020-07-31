import json, requests, time
import numpy as np
import pandas as pd
from pandas import ExcelWriter

url = "http://fanselect.net:8079/FSWebService"
user_ws, pass_ws = 'XXX', 'XXX'
power_factor = 1.04

# Get all the possible fans
all_results = False

def fan_ws(request_string, url):	
	ws_output = requests.post(url=url, data=request_string)
	return ws_output

def get_response(dict_request):
	dict_json = json.dumps(dict_request)
	url_response = fan_ws(dict_json, url)
	url_result = json.loads(url_response.text)
	return url_result

def sort_function(lst, n):
	lst.sort(key = lambda x: x[n])
	return lst 

# Get SessionID
session_dict = {
	'cmd': 'create_session',
	'username': user_ws,
	'password': pass_ws
}

session_id = get_response(session_dict)['SESSIONID']
print('Session ID:', session_id)
print('\n')

# Pandas import
# Open the quotation file
excel_file = 'EC_FANS.xlsx'
df = pd.read_excel(excel_file, usecols= ['Article no', 'Description', 'ID', 'Gross price', 'Plate'],
	dtype={'Article no': str, 'Gross price': float, 'Plate': float})

print('List of fans:')
print(df.head())
print('\n')

# Open the quotation file
excel_file = 'Fans per AHU.xlsx'
df_data = pd.read_excel(excel_file)
# We need to know the W, not the kW
df_data['Consump. kW'] = 1000*df_data['Consump. kW']

# AHU size for merge
excel_file = 'AHU_SIZE.xlsx'
df_size = pd.read_excel(excel_file)

# Merge operation
df_data = pd.merge(df_data, df_size, on='AHU')
cols = ['Height', 'Width', 'Airflow', 'Static Press.']

for col in cols:
	df_data[col] = df_data[col].astype(str)

print('Data input:')
print(df_data.head())
print('\n')

inner_list, outter_list = [], []

# Check execution time
start_time = time.time()

for j in range(len(df_data['Line'])):
	line = df_data['Line'].iloc[j]
	ref = df_data['Ref'].iloc[j]
	ahu = df_data['AHU'].iloc[j]
	height = df_data['Height'].iloc[j]
	width = df_data['Width'].iloc[j]
	qv = df_data['Airflow'].iloc[j]
	psf = df_data['Static Press.'].iloc[j]

	# Old fan data
	old_id = df_data['ID'].iloc[j]
	old_fan = df.loc[df['ID'] == old_id, 'Description'].values[0]
	old_article_no = df.loc[df['ID'] == old_id, 'Article no'].values[0]
	old_no_fans = df_data['No Fans'].iloc[j]
	old_consump = df_data['Consump. kW'].iloc[j]	
	old_cost = df_data['Gross price'].iloc[j]

	file_name = df_data['File name'].iloc[j]

	# Create a new filter for fans df to speed up the process
	df['Fan rows'] = np.floor(int(height)/df['Plate'])
	df_fans = df.loc[df['Fan rows'] > 0, :]
	df_fans['Original fans'] = old_no_fans
	df_fans['Fan columns'] = np.floor(int(width)/df_fans['Plate'])
	df_fans['Max fans dim'] = df_fans['Fan rows'] * df_fans['Fan columns']
	df_fans['Max fans price'] = np.floor(old_cost/df_fans['Gross price'])
	df_fans['Max fans'] = df_fans[['Max fans dim', 'Max fans price', 'Original fans']].min(axis=1)

	time.sleep(1)

	# Loop for fans on each number of line
	for i in range(len(df_fans['Article no'])):
		max_array = int(df_fans['Max fans'].iloc[i]) + 1 # We always need the +1
		# Check several fan configuration
		for n in range(1, max_array):

			# Set values
			new_article_no = df_fans['Article no'].iloc[i]
			new_cost = df_fans['Gross price'].iloc[i]
			print('File name:', file_name)
			print('Line:', line)

			# Fan request - key point
			fan_dict = {
				'language': 'EN',
				'unit_system': 'm',
				'username': user_ws,
				'password': pass_ws,
				'cmd': 'select',
				'cmd_param': '0',
				'zawall_mode': 'ZAWALL_PLUS',
				'zawall_size': n,
				'qv': qv,
				'psf': psf,
				'spec_products': 'PF_00',
				'article_no': new_article_no,
				#'current_phase': '3',
				#'voltage': '400',
				'nominal_frequency': '50',
				'installation_height_mm': height,
				'installation_width_mm': width,
				'installation_length_mm': '2000',
				'installation_mode': 'RLT_2017',
				'sessionid': session_id
			}

			print(fan_dict)
			print('\n')

			try:
				#no_fans = get_response(fan_dict)['ZAWALL_SIZE']
				new_consump = get_response(fan_dict)['ZA_PSYS']

				if new_consump <= (old_consump*power_factor):
					new_fan = df.loc[df['Article no'] == new_article_no, 'Description'].values[0]
					zawall_arr = get_response(fan_dict)['ZAWALL_ARRANGEMENT']
					new_no_fans = 1 if zawall_arr == 0 else int(zawall_arr[:2])
					new_cost = new_no_fans*new_cost

					print('Number of line:', line)
					print('AHU:', ahu)
					print('New fan:', new_fan)
					print('New article no:', new_article_no)
					print('New number of fans:', new_no_fans)
					print('Power input W:', new_consump)
					print('New cost:', new_cost)		
					print('\n')

					inner_list.append([line, ref, ahu, height, width, qv, psf,
						old_fan, old_article_no, old_no_fans, old_consump, old_cost,
						new_fan, new_article_no, new_no_fans, new_consump, new_cost, file_name])

					# Stop the loop
					print('Loop stopping!')
					print('\n')
					break

			except:
				pass
	
	print("--- %s seconds ---" % (time.time() - start_time))
	print('\n')

	if len(inner_list) > 2:
		inner_list = sort_function(inner_list, len(inner_list[0])-2)

	print('Inner list - sorted:')
	print(inner_list)
	print('\n')	

	# Once checked all the items and gathered the entire list, get the cheapest one only if all_results applies
	inner_len = len(inner_list)

	try:
		if all_results:
			for row in range(inner_len):
				outter_list.append(inner_list[row])
		else:
			outter_list.append(inner_list[0])
	except:
		print('----------')
		print('ERROR NOT FAN FOUND')
		print('----------')
		print('\n')

	inner_list = []

# Save all the results to a new dataframe
col = ['Line', 'Ref', 'AHU', 'Height', 'Width', 'Airflow', 'Static Press.',
		'Old fan', 'Old article no', 'Old no fans', 'Old consump.', 'Old cost',
		'New fan', 'New article no', 'New no fans', 'New consump.', 'New cost', 'Filename']

result = pd.DataFrame(outter_list, columns = col)
result['Savings'] = result['Old cost'] - result['New cost']
print(result)

# Export to Excel
name = 'Fans Results_new.xlsx'
writer = pd.ExcelWriter(name)
result.to_excel(writer, index = False)
writer.save()

# Stats
print('\n')
total_old = result['Old cost'].sum()
total_new = result['New cost'].sum()
savings = total_old - total_new
per = savings/total_old
print('Total savings:', savings)
print('Percentage:', per)
print('\n')