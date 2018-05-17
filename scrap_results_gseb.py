import csv
import os
import re
import sys
from bs4 import BeautifulSoup
from time import sleep
import logging
import logging.handlers
from datetime import datetime
from urllib import request


# logging
LOG_FILENAME = "scrap_results_info_gseb.out"

log = logging.getLogger('ScrapResultsInfoLogger')
log.setLevel(logging.DEBUG)

# Add the log message handler to the logger
handler = logging.handlers.RotatingFileHandler(
              LOG_FILENAME, maxBytes=10*1024*1024, backupCount=5)

log.addHandler(handler)

log.info('\n\n-----------------------------------------------------')
log.info('Time of Execution: {}'.format(datetime.now()))
log.info('-----------------------------------------------------\n\n')

# URL = "http://www.gseb.org/522lacigam/sci/B18/67/B186739.html"
URL = 'http://www.gseb.org/522lacigam/sci/{}/{}/{}.html'

bad_chars = r'\xa0\t\n\\'
rgx = re.compile('[%s]' % bad_chars)


# scrap result info from url with roll no. range from l to h
# Note: Lot of things in here are hard coded. This isn't ideal at all.
def scrap_results(series, l, h):

    assert(len(series) == 1)

    try:
        FILENAME = "{}-{}.csv".format(l, h)

        log.info("open: {}".format(FILENAME))

        file = open(FILENAME, 'w')
        fieldnames = [  'Seat No',
                        'Name',
                        'Group',
                        'Percentile',
                        'Science Percentile',
                        'Theory Percentile',
                        'Result',
                        'Total Marks',
                        'Total Obtained Marks']
        writer = csv.writer(file)

        writer.writerow(fieldnames)

        for sid in range(l, h):

            sid = str(sid)

            log.info("Getting result for sid: {}".format(sid))
            print("Getting result for sid: {}".format(sid))

            url = URL.format(series + sid[0:2], sid[2:4], series + sid)

            # should check for error in request.urlopen
            soup = BeautifulSoup(request.urlopen(url).read(), 'html.parser')

            tables = soup.find_all('table')

            rows = tables[0].find_all('tr')

            # find all table data in the rows. Also, clean it.
            temp = [[rgx.sub('', td.text) for td in tr.findAll("td")] for tr in rows]

            # Above opeartion will return list of list of string. So convert it into list of string
            data = [d[0] for d in temp]

            # Just hard coding it to obtain the respective fields
            # Note that this works becuase the fields are static.
            writer.writerow([data[0][9:16],
                            data[0][22:],
                            data[1][8],
                            data[1][-2:],
                            data[2][21:23],
                            data[2][-3:-1],
                            data[3][9:],
                            data[13][11:14],
                            data[13][14:17]])

            # slow down a bit
            sleep(0.005)

    except KeyboardInterrupt:
        log.info("Keyboard Interrupt.")
        sys.exit(0)

    except Exception as e:
        log.error("[!] Error Occurred: {}".format(e))
        sys.exit(1)

    finally:
        log.info("------------------------------------------------------------\n")
        if file:
            file.close()


if __name__ == '__main__':

    series = 'B'
    low = 100001
    high = 100010
    step = 10000

    for i in range(low, high, step):
        l = i

        near_end = high - (high % step)
        if (i >= near_end):
            h = high
        else:
            h = i + step

        log.info("-------------------------------\n")
        log.info("l = {}\nh = {}\n".format(l, h))
        scrap_results(series, l, h)

