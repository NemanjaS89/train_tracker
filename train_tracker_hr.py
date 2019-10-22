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
browser.get('http://vred.hzinfra.hr/hzinfo/Default.asp?Category=hzinfo&Service=TPVG&SCREEN=1')

for i in range (sheet.nrows):
    
    text_field = browser.find_element_by_name('VAG')
    
    text_field.send_keys(str(sheet.cell_value(i, 0)))
    
    submit_button = browser.find_element_by_xpath("//input[@type='SUBMIT']")
    submit_button.click()
    
    soup = BeautifulSoup(browser.page_source, 'lxml')
    title = soup.find_all('title')[0]
    
    if title.text == "Trenutno stanje vagona":
        status = soup.find_all('td')[3]
        location = soup.find_all('td')[4]
        time = soup.find_all('td')[2]
            
        worksheet.write(row, col, str(sheet.cell_value(i, 0)))
        worksheet.write(row, col + 1, status.text)
        worksheet.write(row, col + 2, location.text)
        worksheet.write(row, col + 3, time.text)
        
        if len(soup.find_all('td')) > 5:
            status_2 = soup.find_all('td')[5]
            worksheet.write(row, col + 2, status_2.text)
    
        back_button = browser.find_element_by_xpath("//input[@type='SUBMIT']")
        back_button.click()
    else:
        status2 = soup.find_all('p')[2]
        worksheet.write(row, col, str(sheet.cell_value(i, 0)))
        worksheet.write(row, col + 1, status2.text)
        submit_button2 = browser.find_element_by_xpath("//input[@type='SUBMIT']")
        submit_button2.click()
        
    row += 1
    
workbook.close()