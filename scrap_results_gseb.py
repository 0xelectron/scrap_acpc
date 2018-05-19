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
        FILENAME = "{}-{}-{}.csv".format(series, l, h)

        log.info("open: {}".format(FILENAME))

        file = open(FILENAME, 'w')

        if series == 'B':
            fieldnames = [
                'Seat No',
                'Name',
                'Group',
                'Percentile',
                'Science Percentile',
                'Theory Percentile',
                'Result',
                'Total Marks',
                'Total Obtained Marks'
            ]
        elif series == 'E':
            fieldnames = [
                'Seat No',
                'Name',
                'Group',
                'Total Obtained Marks',
                'Total Marks',
                'Percentile Rank'
            ]
        else:
            log.error("[!] Invalid Series.")
            print("[!] Invalid Series")
            sys.exit(0)

        writer = csv.writer(file)

        writer.writerow(fieldnames)

        for sid in range(l, h + 1):

            sid = str(sid)

            log.info("Getting result for sid: {}".format(sid))
            print("Getting result for sid: {}".format(sid))

            url = URL.format(series + sid[0:2], sid[2:4], series + sid)

            try:
                html_data = request.urlopen(url).read()
            except request.HTTPError as e:
                print(e)
                log.error("[!] Student Id: {}, Error: {}".format(sid, e))
                continue

            # parse the html data
            soup = BeautifulSoup(html_data, 'html.parser')

            tables = soup.find_all('table')

            rows = tables[0].find_all('tr')

            # find all table data in the rows. Also, clean it.
            temp = [[rgx.sub('', td.text) for td in tr.findAll("td")] for tr in rows]

            # Above opeartion will return list of list of string. So convert it into list of string
            data = [d[0] for d in temp]

            if series == 'B':
                # find the string containing marks
                marks = [d for d in data if 'Total Marks'.lower() in d.lower()][1]

                # index of first digit in marks substring
                ifd = re.search('\d', marks).start()

                if ('--' in data[1]):
                    percentile = '--'
                else:
                    percentile = data[1][-5:]

                # Just hard coding it to obtain the respective fields
                # Note that this works becuase the fields are static.
                writer.writerow([data[0][9:16],
                                data[0][22:],
                                data[1][8],
                                percentile,
                                data[2][21:23],
                                data[2][-3:-1],
                                data[3][9:],
                                marks[ifd: ifd + 3],
                                marks[ifd + 3: ifd + 6]])
            elif series == 'E':

                writer.writerow([
                    data[0][-7:],
                    data[1],
                    data[2][-1],
                    data[6][-6:],
                    data[7][-6:],
                    data[8][-6:]
                ])

            else:
                log.error("[!] Invalid Series.")
                print("[!] Invalid Series")
                sys.exit(0)

            # slow down a bit
            sleep(0.005)

    except KeyboardInterrupt:
        log.info("Keyboard Interrupt.")
        print("Keyboard Interrupt")
        sys.exit(0)

    except Exception as e:
        log.error("[!] Error Occurred: {}".format(e))
        sys.exit(1)

    finally:
        log.info("------------------------------------------------------------\n")
        if file:
            file.close()


if __name__ == '__main__':

    series = 'E'
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
        log.info("Series = {}\nl = {}\nh = {}\n".format(series, l, h))
        scrap_results(series, l, h)

