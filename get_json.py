import requests
import os
from collections import OrderedDict
import simplejson as json
import xlrd
import pickle

#import pan

# General
WORK_DIR = r'C:\temp\2019-06-27' 
JSON_NAME = 'data_json.json'

# Manora
NAME_MENORA = 'Menora'
MANORA_URL = 'https://www.menoramivt.co.il/wps/wcm/connect/a29b9bd0-2daa-45b8-a067-9db0867fe03f/%D7%90%D7%95%D7%9E%D7%92%D7%94+21.06.2019.xls?MOD=AJPERES'
MANORA_XLS_NAME = 'manora.xls'
ROW_START_TABLE_MANORA = 6

# Altoler
NAME_ALTOLER = 'Altshuler Shaham'
ALTOLER_URLS = ['https://www.as-invest.co.il/media/10030/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%90%D7%A4%D7%A8%D7%99%D7%9C-2019.xlsx','https://www.as-invest.co.il/media/9799/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%9E%D7%A8%D7%A5-2019.xlsx','https://www.as-invest.co.il/media/9551/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%A4%D7%91%D7%A8%D7%95%D7%90%D7%A8-2019.xlsx','https://www.as-invest.co.il/media/9414/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%99%D7%A0%D7%95%D7%90%D7%A8-2019.xlsx','https://www.as-invest.co.il/media/9413/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%93%D7%A6%D7%9E%D7%91%D7%A8-2018.xlsx', 'https://www.as-invest.co.il/media/9168/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%A0%D7%95%D7%91%D7%9E%D7%91%D7%A8-2018.xlsx','https://www.as-invest.co.il/media/8900/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%90%D7%95%D7%A7%D7%98%D7%95%D7%91%D7%A8-2018.xlsx','https://www.as-invest.co.il/media/8676/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%A1%D7%A4%D7%98%D7%9E%D7%91%D7%A8-2018.xlsx','https://www.as-invest.co.il/media/8675/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%90%D7%95%D7%92%D7%95%D7%A1%D7%98-2018.xlsx','https://www.as-invest.co.il/media/8599/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%99%D7%95%D7%9C%D7%99-2018.xlsx','https://www.as-invest.co.il/media/8235/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%99%D7%95%D7%A0%D7%99-2018-1.xlsx','https://www.as-invest.co.il/media/8051/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%9E%D7%90%D7%99-2018.xlsx','https://www.as-invest.co.il/media/8050/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%90%D7%A4%D7%A8%D7%99%D7%9C-2018.xlsx','https://www.as-invest.co.il/media/7615/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%97%D7%95%D7%93%D7%A9-%D7%9E%D7%A8%D7%A5-2018.xlsx','https://www.as-invest.co.il/media/7614/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%97%D7%95%D7%93%D7%A9-%D7%A4%D7%91%D7%A8%D7%95%D7%90%D7%A8-2018.xlsx','https://www.as-invest.co.il/media/7232/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%99%D7%A0%D7%95%D7%90%D7%A8-2018.xlsx']
ALTOLER_XLS_NAME = 'altoler.xls'
ROW_START_TABLE_ALTOLER = 3

# Psagot
NAME_PSAGOT = 'Psagot'
PSAGOT_URLS = ['https://www.psagot.co.il/heb/PensionSavings/Documents/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA%20%D7%9B%D7%9C%D7%9C%D7%99%D7%95%D7%AA/2019/%D7%90%D7%A1%D7%99%D7%A4%D7%94%20%D7%9B%D7%9C%D7%9C%D7%99%D7%AA%20%D7%92%D7%9E%D7%9C%201.1.19-19.6.19.xlsx', 'https://www.psagot.co.il/heb/PensionSavings/Documents/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA%20%D7%9B%D7%9C%D7%9C%D7%99%D7%95%D7%AA/2019/%D7%90%D7%A1%D7%99%D7%A4%D7%94%20%D7%9B%D7%9C%D7%9C%D7%99%D7%AA%202018%20%D7%92%D7%9E%D7%9C%2001.01.18-31.12.18.xls']
PSAGOT_XLS_NAME = 'psagot.xls'
ROW_START_TABLE_PSAGOT = 2

# Fileds 
PENSION_COMPANY = 'pension_company'
COMPANY_NAME = 'company_name' 
DATE = 'date'
TOPICS = 'topics'
ENTROPY_RECOMMENDATION = 'entropy_recommendation'
PENSION_COMPANY_VOTE = 'pension_company_vote'
FINAL_VERDICT = 'final_verdict'
NEEDED_MAJORITY = 'needed_majority'
CONFLICT_OF_INTEREST = 'conflict_of_interest'

def get_last_recored(current, value ,is_time):
    if type(value) == str and len(value) > 0 or type(value) == float:
    	if is_time:
    		current = str(xlrd.xldate.xldate_as_datetime(value, wb_manora.datemode))
    	else:
    		current = value
    return current

def get_data_sheet(url, work_dir, xls_name):
	manora_data = requests.get(url)
	with open(os.path.join(work_dir, xls_name), 'wb') as f:
		f.write(manora_data.content)
	wb = xlrd.open_workbook(os.path.join(work_dir, xls_name))
	sh = wb.sheet_by_index(0)
	return sh, wb

data_base = []

sh_manora, wb_manora = get_data_sheet(MANORA_URL, WORK_DIR, MANORA_XLS_NAME)
current_company_name, current_data = '', ''
for rownum in range(ROW_START_TABLE_MANORA, sh_manora.nrows - 9):
	row_values = sh_manora.row_values(rownum)
	data = OrderedDict()
	current_company_name = get_last_recored(current_company_name, row_values[0], False)
	current_data = get_last_recored(current_data, row_values[3], True)
	data[PENSION_COMPANY] = NAME_ALTOLER
	data[COMPANY_NAME] = current_company_name
	data[DATE] = current_data
	data[TOPICS] = row_values[5]
	data[ENTROPY_RECOMMENDATION] = row_values[6]
	data[PENSION_COMPANY_VOTE] = row_values[7]
	data[FINAL_VERDICT] = row_values[8]
	data[NEEDED_MAJORITY] = row_values[9]
	data[CONFLICT_OF_INTEREST] = row_values[13]
	j = json.dumps(data)
	data_base.append(j)

for psagot_url in PSAGOT_URLS:
	sh_psagot, wb_psagot = get_data_sheet(psagot_url, WORK_DIR, PSAGOT_XLS_NAME)
	for rownum in range(ROW_START_TABLE_PSAGOT, sh_psagot.nrows - 2):
		row_values = sh_psagot.row_values(rownum)
		data = OrderedDict()
		data[PENSION_COMPANY] = NAME_PSAGOT
		data[COMPANY_NAME] = row_values[1]
		data[DATE] = str(xlrd.xldate.xldate_as_datetime(row_values[5], wb_psagot.datemode))
		data[TOPICS] = row_values[7]
		data[ENTROPY_RECOMMENDATION] = row_values[8] if len(row_values[15]) > 0 else "not " + row_values[8]
		data[PENSION_COMPANY_VOTE] = row_values[8]
		data[FINAL_VERDICT] = row_values[9]
		data[NEEDED_MAJORITY] = row_values[10]
		data[CONFLICT_OF_INTEREST] = row_values[14]
		j = json.dumps(data)
		data_base.append(j)

for altoler_url in ALTOLER_URLS:
	sh_altoler, wb_altoler = get_data_sheet(altoler_url, WORK_DIR, ALTOLER_XLS_NAME)
	for rownum in range(ROW_START_TABLE_ALTOLER, sh_altoler.nrows - 3):
		row_values = sh_altoler.row_values(rownum)
		data = OrderedDict()
		current_company_name = get_last_recored(current_company_name, row_values[1], False)
		current_data = get_last_recored(current_data, row_values[7], False)
		data[PENSION_COMPANY] = NAME_MENORA
		data[COMPANY_NAME] = current_company_name
		data[DATE] = current_data
		data[TOPICS] = row_values[8]
		data[ENTROPY_RECOMMENDATION] = "No Recommendation"
		data[PENSION_COMPANY_VOTE] = row_values[9]
		data[FINAL_VERDICT] = row_values[10]
		data[NEEDED_MAJORITY] = row_values[11]
		data[CONFLICT_OF_INTEREST] = row_values[14]
		j = json.dumps(data)
		data_base.append(j)

with open(os.path.join(WORK_DIR, JSON_NAME), 'wb') as f:
    pickle.dump(data_base, f)