import xlsxwriter
import os
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
import logging
import logging.handlers
from datetime import datetime


# logging
LOG_FILENAME = "scrap_merit_info_acpdc.out"

my_logger = logging.getLogger('ScrapMeritInfoLogger')
my_logger.setLevel(logging.DEBUG)

# Add the log message handler to the logger
handler = logging.handlers.RotatingFileHandler(
              LOG_FILENAME, maxBytes=10*1024*1024, backupCount=5)

my_logger.addHandler(handler)

my_logger.info('\n\n-----------------------------------------------------')
my_logger.info('Time of Execution: {}'.format(datetime.now()))
my_logger.info('-----------------------------------------------------\n\n')

URL = "http://acpdc.in/result/result_search.asp"

# FILENAMES = []

bad_chars = r'\xa0\t\n\\'
rgx = re.compile('[%s]' % bad_chars)


# scrap merit info from url with roll no. range from l to h
def main(l, h):

    try:
        FILENAME = "{}-{}.xlsx".format(l, h)

        my_logger.info("open(out_file): {}".format(FILENAME))
        out_file = xlsxwriter.Workbook(FILENAME)
        sheet = out_file.add_worksheet()

        row = 1
        col = 0

        sheet.write('A1', 'Merit No')
        sheet.write('B1', 'Roll No')
        sheet.write('C1', 'Name')
        sheet.write('D1', 'Allotted Institute Name')
        sheet.write('E1', 'Allotted Course Name')
        sheet.write('F1', 'Allotted Category')
        sheet.write('G1', 'Basic Category')
        sheet.write('H1', 'Allotted Status')


        driver = webdriver.Chrome()
        driver.get(URL)

        for gid in range(l, h):

            my_logger.info("Getting merit info for gid: {}".format(gid))

            elm = driver.find_element_by_name('inno')
            elm.clear()
            elm.send_keys(str(gid) + Keys.RETURN)

            soup = BeautifulSoup(driver.page_source, 'html.parser')

            tables = soup.find_all('table')
            if not tables:
                driver.back()
                continue

            rows = tables[0].find_all('tr')
            data = [[rgx.sub('', td.text) for td in tr.findAll("td")] for tr in rows]

            sheet.write(row, 0, data[1][1])
            sheet.write(row, 1, data[1][3])
            sheet.write(row, 2, data[2][1])
            sheet.write(row, 3, data[3][1])
            sheet.write(row, 4, data[4][1])
            sheet.write(row, 5, data[5][1])
            sheet.write(row, 6, data[5][3])
            sheet.write(row, 7, data[6][1])

            row += 1

            driver.back()
            sleep(0.05)

    except KeyboardInterrupt:
        sys.exit(0)

    finally:
        my_logger.info("------------------------------------------------------------\n")
        if out_file:
            out_file.close()
        if driver:
            driver.close()


if __name__ == '__main__':
    for i in range(1000000, 1039490, 10000):
        l = i 
        if (i == 1030000):
            h = 1039491
        else:
            h = i + 10000
        my_logger.info("-------------------------------\n")
        my_logger.info("l = {}\nh = {}\n".format(l, h))
        main(l, h)
