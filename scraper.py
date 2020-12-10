from bs4 import BeautifulSoup as bs
import sys
import requests
import re
import pandas as pd
import time
import os
from datetime import date

#Disable Web Scraping Warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
session = requests.Session()


url = 'https://br.investing.com/indices/bovespa-components'
headers = {"User-Agent": "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.2.8) Gecko/20100722 Firefox/3.6.8 GTB7.1 (.NET CLR 3.5.30729)", "Referer": "http://example.com"}


response = session.get(url, verify=False, headers=headers)
soup = bs(response.text,'lxml')


names = [child.find('td',class_='bold left noWrap elp plusIconTd').a.text for child in soup.tbody.find_all('tr')]
last = [float(child.find_all('td')[2].text.replace(',','.')) for child in soup.tbody.find_all('tr')]
_max = [float(child.find_all('td')[3].text.replace(',','.')) for child in soup.tbody.find_all('tr')]
_min = [float(child.find_all('td')[4].text.replace(',','.')) for child in soup.tbody.find_all('tr')]
var = [float(child.find_all('td')[5].text.replace(',','.')) for child in soup.tbody.find_all('tr')]
var_pct =[float(child.find_all('td')[6].text.replace(',','.').replace('%',''))/100 for child in soup.tbody.find_all('tr')]
vol = [(float(child.find_all('td')[7].text.replace(',','.').replace('M','')) * 1000000) if re.search('M',child.find_all('td')[7].text) else (float(child.find_all('td')[7].text.replace('K','').replace(',','.')) * 1000) if re.search('K',child.find_all('td')[7].text) else 'Unknown' for child in soup.tbody.find_all('tr')]

df = pd.DataFrame({'Nome': names, 'Ultimo': last, 'Maxima': _max, 'Minima': _min, 'Var.': var, 'Var.%': var_pct, 'Vol.': vol})

writer = pd.ExcelWriter(os.path.dirname(os.path.abspath(sys.argv[0])) + '\\' + str((date.today().strftime("%d-%m-%Y"))) + '.xlsx', engine = 'xlsxwriter',options={'strings_to_numbers': True})

df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)

workbook = writer.book
worksheet = writer.sheets['Sheet1']

format1 = workbook.add_format({'num_format': '0.00'})
format2 = workbook.add_format({'num_format': '0.00%'})
format3 = workbook.add_format({'align': 'left'})

worksheet.set_column('A:A', 22, format3)
worksheet.set_column('B:B', 9, format1)
worksheet.set_column('C:C', 9, format1)
worksheet.set_column('D:D', 9, format1)
worksheet.set_column('E:E', 9, format1)
worksheet.set_column('F:F', 9, format2)
worksheet.set_column('G:G', 10)

worksheet.freeze_panes(1, 0)

writer.save()
