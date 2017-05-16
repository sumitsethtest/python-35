import sys
import re
import urllib2
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import logging
import time
import traceback
import requests
from time import sleep
from selenium import webdriver

match_text=re.compile('Domain Name|Domain|licensed|ftpquota',re.IGNORECASE)
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

test_log = "test" + time.strftime("%Y%m%d-%H%M%S") + ".log"
logging.basicConfig(filename=test_log,level=logging.INFO)

results=[]
driver = webdriver.Ie("C:\Python27\IEDriverServer.exe")

logging.info("INITIATING WEB SCRAPING")
html_string = 'BLANK'

def scrape(url):
	global driver
	page_url = url
	driver.get(page_url)
	html=driver.page_source
	soup = BeautifulSoup(html,'html.parser')
	names=soup.find_all("h3", {"class": "agent-name"})
	phones=soup.find_all("div", {"class":"agent-phone hidden-xs hidden-sm"})
	addresses=soup.find_all("div", {"class":"agent-address"})
	cities=soup.find_all("span", {"class":"c-address-city"})
	postalcodes=soup.find_all("span", {"class":"c-address-postal-code"})
	emails=soup.find_all('a', {'class': 'social-link social-link-email visible-xs visible-sm'})
	#alldetails=zip(names,phones,addresses,cities,postalcodes,emails)
	for i in range(len(names)):
		print(names[i].text)
		print(phones[i].text)
		print(addresses[i].text)
		print(cities[i].text)
		print(postalcodes[i].text)
		print(emails[i]['href'])
		print("########################################")





try:
	scrape('https://agents.allstate.com/usa/wa/seattle')
	driver.close()
	driver.quit()
except:
	logging.error(sys.exc_info()[1])	
	driver.close()
	driver.quit()

