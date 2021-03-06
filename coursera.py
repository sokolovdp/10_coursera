import argparse
import sys
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import requests
from bs4 import BeautifulSoup
import pandas as pd

coursera_page = "https://www.coursera.org/sitemap~www~courses.xml"


def get_html_from_url(url_full: "str") -> "dict":
    response = requests.get(url_full)
    if response.ok:
        return dict(html=response.text, url=response.url, err=None)
    else:
        return dict(html=None, url=None, err=response.status_code)


def get_courses_list():
    page = get_html_from_url(coursera_page)
    if page['html']:
        soup = BeautifulSoup(page['html'], "lxml")
        return [url_full.text for url_full in soup.find_all('loc')]


def get_course_html(course_url: "str") -> "str":
    page = get_html_from_url(course_url)
    if page['html']:
        return page['html']
    else:
        print("can't load course page {}, error {}, info ignored".format(course_url, page['err']))


def parse_course_html(html: "str") -> "dict":
    soup = BeautifulSoup(html, "lxml")
    title = soup.find('meta', attrs={'property': 'og:title'}).attrs['content'].replace(" | Coursera", "")
    start_date = soup.find('div', attrs={'class': 'startdate rc-StartDateString caption-text'}).text
    language = soup.find('div', attrs={'class': 'rc-Language'}).text
    duration = len(soup.find_all('div', attrs={'class': 'week-heading body-2-text'}))
    rating = soup.find_all('div', attrs={'class': 'ratings-text headline-2-text'})
    if rating:  # check if rating is not empty [list]
        rating = rating[0].contents[0].text
    else:
        rating = None
    return dict(Title=title, Date=start_date, Language=language, Weeks=duration, Rating=rating)


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
    courses_html = [get_course_html(url) for url in courses_urls[:number_courses]]
    courses_info = [parse_course_html(html) for html in courses_html if html]
    output_courses_info_to_xlsx(courses_info, filename)
    print("info for {} courses saved in {}".format(number_courses, filename))


if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description='program saves info about N first courses from Coursera feed to excel FILE')
    ap.add_argument("--n", dest="n", action="store", type=int, default=20, help="  number of courses")
    ap.add_argument("--file", dest="file", action="store", default='courses_info.xlsx', help="  file name")
    args = ap.parse_args(sys.argv[1:])

    main(args.file, args.n)
