import PyPDF2, glob, os, csv, sys
from PyPDF2 import PdfReader
from pandas import ExcelWriter
import pandas as pd

# Standard or US ------------------------------------------
standard_version = True

# Get parameters database ---------------------------------
excel_file = 'PARAMETERS.xlsx'
df_param = pd.read_excel(excel_file)

# Function to extract the pageContent ---------------------
def extractContent(pageNumber):
	pageObj = reader.pages[pageNumber]
	pageContent = pageObj.extract_text()

	# Create unique identifiers for extraction
	pageContent = pageContent.replace("\n","")
	return pageContent
# ---------------------------------------------------------

# Get SINGLE value function---------------------------------
def get_value_function(pageContent, wordStart, wordEnd, min_len = 1, max_len = 45):
	# wordStart and wordEnd are unique values, not lists
	posStart = pageContent.index(wordStart) + len(wordStart)
	newContent = pageContent[posStart:]

	#posEnd = indexFunction(wordEnd, newContent)
	posEnd = newContent.index(wordEnd)
	unitFeature = newContent[:posEnd].strip()

	# Check the lenght in order to avoid errors
	if (len(unitFeature) > min_len) & (len(unitFeature) < max_len):
		return unitFeature
	else:
		print('Not valid feature. Content lenght is not ok!')
		return 'Error flag!'
# ---------------------------------------------------------

# Function to get the SystemairCAD version ----------------
def syscad_ver():
	# Extract only first page
	pageContent = extractContent(0)

	# Get version
	wordStart, wordEnd = 'SystemairCAD 2.0 Geniox-1/', ' |'
	try:
		global syscad_version
		syscad_version = get_value_function(pageContent, wordStart, wordEnd, 0)
		print(f"SystemairCAD version: {syscad_version}\n")
	except:
		print('Please check first page of PDF. It should be "SystemairCAD 2.0 Geniox-1"\n')
# ---------------------------------------------------------

# Function to get the range of pages of each unit ---------
def pagesFunction():
	# Array format - starting at zero (0)
	aPageStart, aPageEnd = [], []
	last_page = number_of_pages - 1

	# Keyword
	keyword = extractWords('pagesFunction', 'keyword')

	# Loop through all the pages
	for pageNumber in range(last_page):
		pageContent = extractContent(pageNumber)

		# Get the number of starting and ending pages
		if keyword in pageContent:
			if len(aPageStart) == 0:
				aPageStart.append(pageNumber)
			else:
				aPageEnd.append(pageNumber - 1)
				aPageStart.append(pageNumber)

	# Get the very last page
	aPageEnd.append(last_page)

	# Human readable format
	hPageStart, hPageEnd = [x+1 for x in aPageStart], [x+1 for x in aPageEnd]
	print(f"Pages (readable format): {hPageStart}, {hPageEnd}\n")

	return aPageStart, aPageEnd
# ---------------------------------------------------------

# Function extractWords -----------------------------------
def extractWords(function, field):
	# Helper to slice the parameters dataset
	wordStart = df_param.loc[(df_param['Version'] == syscad_version) &
	(df_param['Function'] == function) &
	(df_param['Field'] == field), 'wordStart'].values[0]

	wordEnd = df_param.loc[(df_param['Version'] == syscad_version) &
	(df_param['Function'] == function) &
	(df_param['Field'] == field), 'wordEnd'].values[0]

	# Check if return one or two words
	if pd.isna(wordEnd):
		return wordStart
	else:
		return wordStart, wordEnd
# ---------------------------------------------------------

# First page function -------------------------------------
def fpFunction():
	print(f"Starting first page function ------------------\n")
	inner_list, outter_list = [], []
	for page, pageEnd in zip(aPageStart, aPageEnd):
		print(f"Looking at page {page+1}")
		pageContent = extractContent(page)

		# Get line
		wordStart, wordEnd = extractWords('fpFunction', 'Line')
		line = get_value_function(pageContent, wordStart, wordEnd)

		# Reset ahu_value for each pageStart
		ahu_value = []

		# Get AHU type
		for ahu in ahus:
			if ahu in pageContent:
				ahu_value.append(ahu)

		# In case of DV10 or DV100, always get the longest one
		if len(ahu_value) == 1:
			ahu = ahu_value[0]
		elif len(ahu_value) > 1:
			print('Possible conflict!')
			final_ahu = ''
			for value in ahu_value:
				if len(value) > len(final_ahu):
					final_ahu = value
			ahu = final_ahu
			print('Final value:', ahu)
			print('\n')
		else:
			ahu = '---'

		# Get reference
		wordStart, wordEnd = extractWords('fpFunction', 'Reference')
		ref = get_value_function(pageContent, wordStart, wordEnd)

		# Airflow
		wordStart, wordEnd = extractWords('fpFunction', 'Airflow')
		airflow = get_value_function(pageContent, wordStart, wordEnd)

		inner_list = [page, pageEnd, line, ahu, ref, airflow]
		outter_list.append(inner_list)

	print(f"\n-----------------------------------------------\n")
	return outter_list
# ---------------------------------------------------------


'''
# Get the range of the pages ------------------------------
syscad_ver() # Mandary always
aPageStart, aPageEnd = pagesFunction()
last_page = aPageEnd[-1:][0]

ahus = ['Geniox 10', 'Geniox 11', 'Geniox 12', 'Geniox 14', 'Geniox 16', 'Geniox 18',
	'Geniox 20', 'Geniox 22', 'Geniox 24', 'Geniox 27', 'Geniox 29', 'Geniox 31', 'Geniox 35',
	'Geniox 38', 'Geniox 41', 'Geniox 44', 'Geniox On 10', 'Geniox On 11', 'Geniox On 12',
	'Geniox On 14', 'Geniox On 16', 'Geniox On 18', 'Geniox On 20', 'Geniox On 22', 'Geniox On 24',
	'Geniox On 27', 'Geniox On 29', 'Geniox On 31']

first_pages = fpFunction()

print(first_pages)

# Test for page content
#print(extractContent(8))
'''


# Possible main -------------------------------------------
def extractFeatures(aWordStart, aWordEnd, pageStart, pageEnd, allowed_pages = 1):
	print(f"Starting extractFeatures function -------------\n")
	outter_list = []
	for page in range(pageStart, pageEnd):
		# Initiate the inner_list and get the page number
		inner_list = []
		
		# Extract page content
		pageContent = extractContent(page)
		print(f"Checking at page number {page+1}")

		for wordStart, wordEnd in zip(aWordStart, aWordEnd):
			print(f"Looking for {wordStart} and {wordEnd}")

			# Work in starting and ending pairs, page by page
			if (wordStart in pageContent) and (wordEnd in pageContent):
				print(f"Found on page {page+1}!\n")
				#print(f"{pageContent}\n")
				unitFeature = get_value_function(pageContent, wordStart, wordEnd)

				# Important in case the next wordStart is above the previous one
				print(f"Feature found: {unitFeature}")
				split_word = unitFeature + wordEnd
				print(f"Split_word: {split_word}")

				# Check if this split_word exists
				if (split_word in pageContent):
					print(f"Split_word in pageContent") 
					posEnd = pageContent.index(split_word)
					pageContent = pageContent[posEnd:]
				else:
					print(f"Split_word {split_word} not found in page {page+1}\n") 
					break

				if unitFeature == 'Error flag!':
					print('Error flag! Length not correct.')
					break
				else:
					inner_list.append(unitFeature)

			# If cheking of additional pages is allowed
			elif allowed_pages > 0:
				if len(inner_list) == 0:
					# Reset inner list
					inner_list = []
					print('No luck this time')
					print('\n')
					# Exit loop and go for the next page
					break
				elif len(inner_list) > 0:
					until_page = page + allowed_pages + 1
					for new_page in range(page + 1, until_page):
						pageContent = extractContent(new_page)
						try:
							print('\n')
							print('New page being checked')
							unitFeature = get_value_function(pageContent, wordStart, wordEnd)
							inner_list.append(unitFeature)
						except:
							print('No luck even in the next page')
							print('\n')
							# Exit loop and go for the next page
							break

		# Check the lenght and append to the outter list
		if len(inner_list) == len(aWordStart):
			print('New entry for the outter list!')
			print('\n')
			# Add the number of page in the inner list
			inner_list = [page + 1, *inner_list] # In order to show real page number
			outter_list.append(inner_list)
			# Reset inner list for next feature
			inner_list = []

	print(f"----------------------------------------------\n")
	try:
		return outter_list
	except:
		print('No outter_list found!')
#----------------------------------------------------------

# Main ----------------------------------------------------
# Open file and read it
path = os.path.dirname(os.path.realpath(__file__))
num_files = len(glob.glob1(path,'*.pdf'))

extList, newList = [], []

# EC_FANS
if standard_version:
	excel_file = 'EC_FANS.xlsx'
else:
	excel_file = 'EC_FANS_US.xlsx'
df_ec = pd.read_excel(excel_file, dtype={'Item': str, 'Gross price': float})

# Empty dataframe
cols = ['Page', 'Airflow', 'Static Press.', 'Motor Power', 'RPM', 'Consump. kW',
		'Line', 'AHU', 'Ref', 'No Fans', 'ID', 'Gross price', 'File name']

df_outter = pd.DataFrame(columns=cols)

for fileName in glob.glob('*.pdf'):
	# Initialize ----------------------------------------------
	ahuSize, ahuLine, aPageStart = [], [], []

	ahus = ['Geniox 10', 'Geniox 11', 'Geniox 12', 'Geniox 14', 'Geniox 16', 'Geniox 18',
	'Geniox 20', 'Geniox 22', 'Geniox 24', 'Geniox 27', 'Geniox 29', 'Geniox 31', 'Geniox 35',
	'Geniox 38', 'Geniox 41', 'Geniox 44', 'Geniox On 10', 'Geniox On 11', 'Geniox On 12',
	'Geniox On 14', 'Geniox On 16', 'Geniox On 18', 'Geniox On 20', 'Geniox On 22', 'Geniox On 24',
	'Geniox On 27', 'Geniox On 29', 'Geniox On 31']

	# Read the pdf --------------------------------------------
	#fileName = fileName[:-4]
	#pdfFileObj = open(fileName + '.pdf', 'rb')
	reader = PdfReader(fileName)

	# Get number of pages -------------------------------------
	number_of_pages = len(reader.pages)
	print(f"File: {fileName} | Number of pages: {number_of_pages}\n")

	# SystemairCAD version ------------------------------------
	syscad_ver() # Mandatory

	# Get the range of the pages ------------------------------
	aPageStart, aPageEnd = pagesFunction()
	last_page = aPageEnd[-1:][0]

	# First page function -------------------------------------
	try:
		units = fpFunction()
	except:
		print('*** Check datasheet language! ***')
		print('\n')
		continue

	columns_units = ['Page start', 'Page end', 'Line', 'AHU', 'Ref', 'Airflow']
	df_units = pd.DataFrame(units, columns = columns_units)

	# To get an accurate reading in the Excel file
	cols = ['Page start', 'Page end']
	for col in cols:
		df_units[col] = df_units[col].astype(int)
		df_units[col] = df_units[col] + 1
	
	#---------------------------------------------------------------------------------------------------------------#
	# Fan scraping

	aWordStart = df_param.loc[(df_param['Version'] == syscad_version) &
	(df_param['Function'] == 'extractFeatures'), 'wordStart'].values.tolist()

	aWordEnd = df_param.loc[(df_param['Version'] == syscad_version) &
	(df_param['Function'] == 'extractFeatures'), 'wordEnd'].values.tolist()


	aWordEnd = [' m', ' Pa', ' kW', ' RPM', ' kW']
	outter = extractFeatures(aWordStart, aWordEnd, 0, last_page, 1)
	columns = ['Page', 'Airflow', 'Static Press.', 'Motor Power', 'RPM', 'Consump. kW']
	df = pd.DataFrame(outter, columns = columns)
	df['Page'] = df['Page'].astype(int)

	#---------------------------------------------------------------------------------------------------------------#
	# Merge the fans with the size of the units
	cols = ['Line', 'AHU', 'Ref']
	line_list, ahu_list, ref_list = [], [], []
	list_of_lists = [line_list, ahu_list, ref_list]

	for fan_page in df['Page'].values:
		for col, list in zip(cols, list_of_lists):
			value = df_units.loc[(df_units['Page start'].values < fan_page) & (df_units['Page end'] > fan_page), col].values[0]
			list.append(value)

	for col, list in zip(cols, list_of_lists):
		df[col] = list
	#---------------------------------------------------------------------------------------------------------------#
	# Dataframe cleaning for several motors
	# Number of fans
	df['No Fans'] = 1
	print('df before fans:')
	print(df['Motor Power'])

	print('\n')
	df.loc[df['Motor Power'].str.contains('total'), 'No Fans'] = df['Motor Power'].str.slice(7, 8, 1)
	df['No Fans'] = df['No Fans'].astype(int)

	# Motor Power
	df.loc[df['Motor Power'].str.contains('nominal'), 'Motor Power'] = df['Motor Power'].str.slice(8)
	df.loc[df['Motor Power'].str.contains('total'), 'Motor Power'] = df['Motor Power'].str.slice(11)

	# ID - cleaning
	cols = ['Motor Power', 'RPM']

	for col in cols:
		df[col] = df[col].str.replace(" ", "")

	df['ID'] = df['Motor Power'] + '-' + df['RPM']

	# Show results
	print('df after fans:')
	print('\n')
	print(df)

	# Old Gross
	df = pd.merge(df, df_ec.loc[:, ['ID', 'Gross price']], on='ID')
	df['Gross price'] = df['Gross price'].values * df['No Fans'].values
	#---------------------------------------------------------------------------------------------------------------#

	if df.empty:
		print('*** Check if unit has metallic or special fans! ***')
		print('\n')
	else:
		# Last tweaks
		df['File name'] = fileName
		df_outter = df_outter.append(df)

# Add the last columns
df_outter = pd.merge(df_outter, df_ec.loc[:, ['ID', 'Item no', 'Description']], on='ID')
df_outter['Total sales price'] = df_outter['Gross price'].values/0.6 # 40% CM

# Reorder columns
new_cols = ['Page', 'Airflow', 'Static Press.', 'Motor Power', 'RPM', 'Consump. kW',
			'Line', 'AHU', 'Ref', 'No Fans', 'ID', 'Item no', 'Description', 'Gross price', 'Total sales price', 'File name']
			
df_outter = df_outter.reindex(columns=new_cols)

# Sort the columns
df_outter = df_outter.sort_values(["File name", "Page"], ascending = (True, True))

# Send to Excel
name = 'Fans per AHU.xlsx'
writer = pd.ExcelWriter(name)
df_outter.to_excel(writer, index = False)
writer.save()	
#pdfFileObj.close()

'''
# Fan version
if standard_version:
	import fan_selection_syscad # Standard
else:
	import fan_selection_syscad_us # US version 230V III
'''