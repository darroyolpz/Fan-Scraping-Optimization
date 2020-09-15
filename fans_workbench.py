import json, requests, time, datetime
from datetime import datetime
import numpy as np
import pandas as pd
from pandas import ExcelWriter

url = "http://fanselect.net:8079/FSWebService"
user_ws, pass_ws = 'xxx', 'xxx'
bluefin_only = True

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

# Open EC fans file
excel_file = 'EC_FANS.xlsx'
df = pd.read_excel(excel_file, usecols= ['Article no', 'Description', 'ID', 'Gross price', 'Plate', 'Bluefin'],
	dtype={'Article no': str, 'Gross price': float, 'Plate': float})

# AHU size for merge
excel_file = 'AHU_SIZE_bench.xlsx'
df_size = pd.read_excel(excel_file)

# Create working dataframe
inner_data = []
for x in range(len(df_size['AHU'])):
	ahu = df_size['AHU'].iloc[x]
	height = df_size['Height'].iloc[x]
	width = df_size['Width'].iloc[x]
	airflow_start = df_size['Airflow start'].iloc[x]
	airflow_end = df_size['Airflow end'].iloc[x]
	step = int(round((airflow_end - airflow_start)/5))
	airflow_end += 1

	# Airflow
	for a in range(airflow_start, airflow_end, step):
		# Static Pressure
		for p in range(200, 1801, 200):
			val = [ahu, height, width, a, p]
			inner_data.append(val)

cols = ['AHU', 'Height', 'Width', 'Airflow', 'Static Press.']
df_data = pd.DataFrame(inner_data, columns = cols)
print(df_data)
print('\n')

# Cleaning for web-service
for col in cols:
	df_data[col] = df_data[col].astype(str)

# Check execution time
start_time = time.time()

# Set empty values for web-service
inner_list = []

for j in range(len(df_data['AHU'])):

	# General data
	ahu = df_data['AHU'].iloc[j]
	height = df_data['Height'].iloc[j]
	width = df_data['Width'].iloc[j]
	qv = df_data['Airflow'].iloc[j]
	psf = df_data['Static Press.'].iloc[j]

	# Create a new filter for fans df to speed up the process
	df['Fan rows'] = np.floor(int(height)/df['Plate'])
	df_fans = df.loc[df['Fan rows'] > 0, :]
	df_fans['Fan columns'] = np.floor(int(width)/df_fans['Plate'])
	df_fans['Max fans dim'] = df_fans['Fan rows'] * df_fans['Fan columns']
	df_fans['Max fans'] = df_fans['Max fans dim']

	# Just in case we have rounding errors
	df_fans = df_fans.loc[df_fans['Max fans'] > 0, :]

	time.sleep(1)

	# Bluefin only?
	if bluefin_only:
		df_fans = df_fans.loc[df_fans['Bluefin'].values == 1, :]
	else:
		df_fans = df_fans.loc[df_fans['Bluefin'].values == 0, :]

	print('Final list of fans to use:')
	print(df_fans)
	print('\n')

	# Loop for fans on each ahu
	for i in range(len(df_fans['Article no'])):
		max_array = int(df_fans['Max fans'].iloc[i]) + 1 # We always need the +1
		# Check several fan configuration
		for n in range(1, max_array):
			# Airflow for several fans - Only for calculation
			try:
				if n > 1:
					qv_calc = str(int(qv)/n)
				else:
					qv_calc = qv
			except:
				qv_calc = qv

			# Set values
			new_article_no = df_fans['Article no'].iloc[i]
			new_cost = df_fans['Gross price'].iloc[i]

			# Fan request - key point
			fan_dict = {
				'language': 'EN',
				'unit_system': 'm',
				'username': user_ws,
				'password': pass_ws,
				'cmd': 'select',
				'cmd_param': '0',
				'qv': qv_calc,
				'psf': psf,
				'spec_products': 'PF_00',
				'article_no': new_article_no,
				'nominal_frequency': '50',
				'sessionid': session_id
			}

			# Check execution time
			start_time = time.time()
			currentDT = datetime.now()
			print (str(currentDT))
			print('** Checking', ahu, 'with', qv, 'm3/h at', psf, 'Pa **')
			print(fan_dict)
			print('\n')

			try:
				new_consump = n*get_response(fan_dict)['ZA_PSYS']
				new_fan = df.loc[df['Article no'] == new_article_no, 'Description'].values[0]
				new_no_fans = n
				new_cost = new_no_fans*new_cost

				print('AHU:', ahu)
				print('New fan:', new_fan)
				print('New article no:', new_article_no)
				print('New number of fans:', new_no_fans)
				print('New consump.:', new_consump)
				print('New cost:', new_cost)		
				print('\n')

				inner_list.append([ahu, height, width, qv, psf,
					new_fan, new_article_no, new_no_fans, new_consump, new_cost])

				# Stop the loop
				print('Loop stopping!')
				print(inner_list)
				print('\n')
				break

			except:
				pass
	
# Send to Excel each AHU
print("--- %s seconds ---" % (time.time() - start_time))
print('\n')

# Save all the results to a new dataframe
col = ['AHU', 'Height', 'Width', 'Airflow', 'Static Press.',
		'New fan', 'New article no', 'New no fans', 'New consump.', 'New cost']

result = pd.DataFrame(inner_list, columns = col)
print(result)

# Export to Excel
name = 'Fans Workbench Results.xlsx'
writer = pd.ExcelWriter(name)
result.to_excel(writer, index = False)
writer.save()
