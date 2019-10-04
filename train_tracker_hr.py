from selenium import webdriver
from bs4 import BeautifulSoup
import xlrd
import xlsxwriter

#initiating the new xlsx
workbook = xlsxwriter.Workbook('trains_hr_generated.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

#reading from the entry xlsx
loc = 'trains_hr.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

#initiating the browser
browser = webdriver.Chrome('resources/chromedriver.exe')
browser.get('http://vred.hzinfra.hr/hzinfo/Default.asp?Category=hzinfo&Service=tpvl&SCREEN=1')

for i in range (sheet.nrows):
    text_field = browser.find_element_by_name('VL')
    
    text_field.send_keys(sheet.cell_value(i, 0))
    submit_button = browser.find_element_by_xpath("//input[@value=' OK ']")
    submit_button.click()
    
    soup = BeautifulSoup(browser.page_source, 'lxml')
    result = soup.find_all('p')[2]
    
    worksheet.write(row, col, sheet.cell_value(i, 0))
    worksheet.write(row, col + 1, result.text)
    row += 1
    
workbook.close()