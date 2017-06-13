import argparse
import sys
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import urllib.request
from bs4 import BeautifulSoup
import pandas as pd


class HTML:
    def __init__(self, html, err):
        self.html = html
        self.err = err
        self.ok = (self.html is not None)


def get_html_from_url(url_full: "str") -> "HTML":
    try:
        response = urllib.request.urlopen(url_full)
        return HTML(response.read(), None)
    except urllib.error.URLError as e1:
        return HTML(None, e1.reason)
    except Exception as e2:
        return HTML(None, e2.reason)


def get_courses_list():
    page = get_html_from_url("https://www.coursera.org/sitemap~www~courses.xml")
    if page.ok:
        soup = BeautifulSoup(page.html, "lxml")
        return [url_full.text for url_full in soup.find_all('loc')]
    else:
        print("can't load list of courses, error {}".format(page.err))
        exit()


def get_course_info(course_url):
    page = get_html_from_url(course_url)
    if page.ok:
        soup = BeautifulSoup(page.html, "lxml")
        title = soup.find('meta', attrs={'property': 'og:title'}).attrs['content'].replace(" | Coursera", "")
        start_date = soup.find('div', attrs={'class': 'startdate rc-StartDateString caption-text'}).text
        language = soup.find('div', attrs={'class': 'rc-Language'}).text
        duration = len(soup.find_all('div', attrs={'class': 'week-heading body-2-text'}))
        rating = soup.find_all('div', attrs={'class': 'ratings-text headline-2-text'})
        if rating:  # check if rating is not empty [list]
            rating = rating[0].contents[0].text
        else:
            rating = None
        return {'1_title': title, '2_date': start_date, '3_language': language, '4_weeks': duration, "5_rating": rating}
    else:
        print("can't load course page {}, error {}, info ignored".format(course_url, page.err))


def output_courses_info_to_xlsx(courses_info, filename):
    book = Workbook()
    sheet = book.active
    sheet.title = "Coursera"
    courses_dataframe = pd.DataFrame(courses_info).fillna('-')  # replace all None values with '-'
    for row in dataframe_to_rows(courses_dataframe, index=True, header=True):
        sheet.append(row)
    book.save(filename=filename)


def main(filename, number_courses):
    courses_urls = get_courses_list()
    print("loaded info about {} courses".format(len(courses_urls)))
    courses_info = [get_course_info(url) for url in courses_urls[:number_courses]]
    output_courses_info_to_xlsx(courses_info, filename)
    print("info for {} courses saved in {}".format(number_courses, filename))


if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description='program saves info about N first courses from Coursera feed to excel FILE')
    ap.add_argument("--n", dest="n", action="store", type=int, default=20, help="  number of courses")
    ap.add_argument("--file", dest="file", action="store", default='courses_info.xlsx', help="  file name")
    args = ap.parse_args(sys.argv[1:])

    main(args.file, args.n)
