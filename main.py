from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.worksheet.table import Table,TableStyleInfo

records = []

r = requests.get('https://www.worldometers.info/world-population/population-by-country/')

c = r.content

soup = BeautifulSoup(c, 'html.parser')

data = soup.find('tbody')
rows = data.find_all('tr')

for index,row in enumerate(rows,1):
    columns = row.find_all('td')

    country = columns[1].text
    population = int(columns[2].text.replace(',',''))
    percentage = columns[3].text

    item = [index,country,population,percentage]

    records.append(item)

records.insert(0, ['N', 'Countries', 'Population','Percentage'])


workbook = Workbook()

file_name = 'Population.xlsx'

workbook.save(file_name)

sheet = workbook['Sheet']
sheet.title = 'Population'
sheet = workbook['Population']

for item in records:
    sheet.append(item)

table = Table(displayName='Population_Data', ref='A1:D236')
style = TableStyleInfo(name = 'TableStyleMedium4', showRowStripes=True, showColumnStripes=True)

table.tableStyleInfo = style
sheet.add_table(table)

font = Font(color='00FF0000', bold=True,italic=True)

for cell_number in range(2,235):
    if sheet[f'C{cell_number}'].value < 5000000:
        sheet[f'C{cell_number}'].font = font

workbook.save(file_name)
workbook.close()















