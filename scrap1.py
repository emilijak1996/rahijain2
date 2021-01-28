import requests
from bs4 import BeautifulSoup
import csv
import xlsxwriter
import json
import xlrd
import os
import time
import math
from datetime import datetime
OUTPUT_DIR	= 'output'
results = []
cities = []
keywords = []
TOT=2
tot=0
file_location="input.xls"
workbook=xlrd.open_workbook(file_location)
sheet=workbook.sheet_by_index(0)

HEADERS = {
			'Content-Type': 'application/json',
			'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0',
			'Origin': 'https://www.justdial.com',
			'Host': 'www.justdial.com',
			'x-cms': 'v2',
			'x-content': 'desktop',
			'x-mp': 'justdial',
			'x-platform': 'web',
			'Accept': 'application/json, text/plain, */*',
			'Connection': 'keep-alive',
			'Cache-Control': 'no-cache, max-age=0, must-revalidate, no-store',
			'Accept-Encoding': 'gzip, deflate, br',
			'TE': 'Trailers'
		}
for row in range(1,sheet.nrows):
	cities=cities+[sheet.cell_value(row,1)]
	print(cities)
for row in range(1,9):
	keywords=keywords+[sheet.cell_value(row,4)]
for keyword in keywords :
	for city in cities :
		tot+=1
		url="https://www.justdial.com/" + city + "/" + keyword
		page = requests.get(url,headers=HEADERS)
		soup = BeautifulSoup(page.text, 'lxml')
		listing_parent=soup.find(class_="lstEmt")
		try:
			listing=listing_parent.find(class_="lng_crcum").text
			total_num=listing.split('+')[0]
		except:
			print("listing don't exist")
		try:
			total_page=math.ceil(float(total_num)/10)
		except:
			total_page=1
		if total_page > 50:
			total_page = 50
		link_parent=soup.find(id="brd_cm_srch")
		try:
			link=link_parent['href']
		except:
			link=url
		try:
			for page in range(1,total_page +1) :
				print(keyword," and ",city, " --",page," page loading...")
				page_link=link + "/page-" +str(page)
				page_data= requests.get(page_link,headers = HEADERS)
				page_data_list= BeautifulSoup(page_data.text, 'lxml')
				data_list=page_data_list.find_all('li', class_='cntanr')
				if len(data_list)==0:
						container = dict()
						container["city"]=city
						container["keyword"] =keyword
						container["store_name"]=""
						container["store_address"]=""
						container["store_phonenumber"]=""
						container["listing"]=0
						results = results +[container]
				else:
					for data in data_list :					
						try:
							store_name = data.find(class_='lng_cont_name').text
						except:
							store_name=""
						try:
							store_address =data.find(class_='cont_fl_addr').text
						except:
							print("Store address don't exist!")
						store_phonenumber=""
						spans=data.find_all('span', class_='mobilesv')
						for span in spans :
							phone_num=span['class'][1]
							if phone_num == "icon-dc" :
								store_phonenumber+="+"
							if phone_num == "icon-fe" :
								store_phonenumber+="("
							if phone_num == "icon-ji" :
								store_phonenumber+="9"
							if phone_num == "icon-yz" :
								store_phonenumber+="1"
							if phone_num == "icon-hg" :
								store_phonenumber+=")"
							if phone_num == "icon-ba" :
								store_phonenumber+="-"
							if phone_num == "icon-yx" :
								store_phonenumber+="2"
							if phone_num == "icon-vu" :
								store_phonenumber+="3"
							if phone_num == "icon-lk" :
								store_phonenumber+="8"
							if phone_num == "icon-po" :
								store_phonenumber+="6"
							if phone_num == "icon-abc" :
								store_phonenumber+="0"
							if phone_num == "icon-nm" :
								store_phonenumber+="7"
							if phone_num == "icon-rq" :
								store_phonenumber+="5"
							if phone_num == "icon-ts" :
								store_phonenumber+="4"
						container = dict()
						container["city"]=city
						container["keyword"] =keyword
						container["store_name"]=store_name
						container["store_address"]=store_address
						container["store_phonenumber"]=store_phonenumber
						container["listing"]=total_num
						results = results +[container]
		except:
			container = dict()
			container["city"]=city
			container["keyword"] =keyword
			container["store_name"]=""
			container["store_address"]=""
			container["store_phonenumber"]=""
			container["listing"]=0
			results = results +[container]
if not os.path.exists(OUTPUT_DIR):
		os.makedirs(OUTPUT_DIR)

	# create output file name based on time
dt = datetime.now().strftime("%d%m%Y%H%M%S")
outputXLSX = os.path.join(OUTPUT_DIR, f"scrape-result-{dt}.xlsx")

workbook = xlsxwriter.Workbook(outputXLSX)
worksheet = workbook.add_worksheet()
	
BASIC_FORMAT = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter'})

worksheet.write("A1", "Search String", BASIC_FORMAT)
worksheet.write("B1", "Search City", BASIC_FORMAT)
worksheet.write("C1", "Store Name", BASIC_FORMAT)
worksheet.write("D1", "Store Address", BASIC_FORMAT)
worksheet.write("E1", "Store Phone Number", BASIC_FORMAT)
worksheet.write("F1", "Listing Numbsers", BASIC_FORMAT)

CELL_WIDTH = 150
worksheet.set_column(0, CELL_WIDTH)
row = 2
print("writting...")
for em in results:
	worksheet.write(f"A{row}", em["keyword"], BASIC_FORMAT)
	worksheet.write(f"B{row}", em["city"], BASIC_FORMAT)
	worksheet.write(f"C{row}", em["store_name"], BASIC_FORMAT)
	worksheet.write(f"D{row}", em["store_address"], BASIC_FORMAT)
	worksheet.write(f"E{row}", em["store_phonenumber"], BASIC_FORMAT)
	worksheet.write(f"F{row}", em["listing"], BASIC_FORMAT)

	
	row += 1

workbook.close()
print(f'+ data written to {outputXLSX}')

