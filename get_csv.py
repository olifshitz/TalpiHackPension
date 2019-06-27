import urllib.request
from bs4 import BeautifulSoup
from selenium import webdriver
import time
import pandas as pd

from my_consts import *

URL = "https://maya.tase.co.il/reports/company?q=%7B%22DateFrom%22:%222018-12-26T22:00:00.000Z%22,%22DateTo%22:%222019-06-26T21:00:00.000Z%22,%22events%22:%5B1500%5D,%22subevents%22:%5B213,1501,1502,1503,1504%5D,%22Page%22:1%7D"

driver = webdriver.Firefox(executable_path=GECKOD_DRIVER_PATH)
driver.get(URL)

report_titles = [element.text for element in driver.find_elements_by_xpath("//*[@class='messageContent ng-binding']")]
report_companies = [element.text for element in driver.find_elements_by_xpath("//*[@class='feedItemCompany ng-scope']")]
report_dates = [element.text for element in driver.find_elements_by_xpath("//*[@class='feedItemDate hidden-xs hidden-sm ng-binding']")]
# report_dates = [element.text for element in driver.find_elements_by_xpath("//*[@class='feedItemDate hidden-md hidden-lg ng-binding']")]

df = pd.DataFrame(list(zip(report_titles, report_companies, report_dates)), columns = ['Title', 'Company', 'Date'])
df.to_csv(RESULT_CSV_PATH, encoding='utf8')