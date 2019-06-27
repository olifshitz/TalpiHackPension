import urllib.request
from bs4 import BeautifulSoup
from selenium import webdriver
import time
import pandas as pd

URL = "https://maya.tase.co.il/reports/company?q=%7B%22DateFrom%22:%222018-12-26T22:00:00.000Z%22,%22DateTo%22:%222019-06-26T21:00:00.000Z%22,%22events%22:%5B1500%5D,%22subevents%22:%5B213,1501,1502,1503,1504%5D,%22Page%22:1%7D"
GECKOD_DRIVER_PATH = r'C:\temp\2019-06-27\geckodriver-v0.24.0-win64\geckodriver.exe'


driver = webdriver.Firefox(executable_path=GECKOD_DRIVER_PATH = r'C:\temp\2019-06-27\geckodriver-v0.24.0-win64\geckodriver.exe')
driver.get(URL)


report_titles = driver.find_elements_by_xpath("//*[@class='feedItemMessage']")
report_companies = driver.find_elements_by_xpath("//*[@class='feedItemCompany ng-scope']")