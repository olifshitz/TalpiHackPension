import requests
import os
from collections import OrderedDict
import simplejson as json
import xlrd

#import pan

# General
WORK_DIR = r'C:\temp\2019-06-27' 
JSON_NAME = 'data_json.json'

# Manora
MANORA_URL_FORMAT = 'https://www.menoramivt.co.il/wps/wcm/connect/a29b9bd0-2daa-45b8-a067-9db0867fe03f/%D7%90%D7%95%D7%9E%D7%92%D7%94+21.06.2019.xls?MOD=AJPERES'
MANORA_XLS_NAME = 'manora.xls'
ROW_START_TABLE_MANORA = 6

# Altoler
ALTOLER_URL_FORMAT = 'https://www.as-invest.co.il/media/10030/%D7%90%D7%A1%D7%99%D7%A4%D7%95%D7%AA-%D7%90%D7%A4%D7%A8%D7%99%D7%9C-2019.xlsx'
ALTOLER_XLS_NAME = 'altoler.xls'

# Fileds 
PENSION_COMPANY = 'pension_company'
COMPANY_NAME = 'company_name' 
DATE = 'date'
TOPICS = 'topics'

def get_last_recored(current, value):
    if type(value) == str and len(value) > 0 or type(value) == float:
    	current = value
    return current

def get_data_shit(url, work_dir, xls_name):
	manora_data = requests.get(url)
	with open(os.path.join(work_dir, xls_name), 'wb') as f:
		f.write(manora_data.content)
	wb = xlrd.open_workbook(os.path.join(work_dir, xls_name))
	sh = wb.sheet_by_index(0)
	return sh

data_base = []

sh_manora = get_data_shit(MANORA_URL_FORMAT, WORK_DIR, MANORA_XLS_NAME)
current_company_name, current_data = '', ''
for rownum in range(ROW_START_TABLE, sh.nrows - 9):
	row_values = sh_manora.row_values(rownum)
	data = OrderedDict()
	current_company_name = get_last_recored(current_company_name, row_values[0])
	current_data = get_last_recored(current_data, row_values[3])
	data[PENSION_COMPANY] = ''
	data[COMPANY_NAME] = current_company_name
	data[DATE] = current_data
	data[TOPICS] = row_values[4]
	data_base.append(data)

'''
sh_manora = get_data_shit(ALTOLER_URL_FORMAT, WORK_DIR, ALTOLER_XLS_NAME)
for rownum in range(ROW_START_TABLE, sh.nrows - 9):
	row_values = sh_manora.row_values(rownum)
	current_company_name = get_last_recored(current_company_name, row_values[0])
	current_data = get_last_recored(current_data, row_values[3])
	data[COMPANY_NAME].append(current_company_name)
	data[DATE].append(current_data)
	data[TOPICS].append(row_values[4])
'''



j = json.dumps(data)
with open(os.path.join(WORK_DIR, JSON_NAME), 'w') as f:
    f.write(j)