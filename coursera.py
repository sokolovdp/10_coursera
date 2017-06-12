from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import urllib.request
from bs4 import BeautifulSoup
import pandas as pd


def get_url_soup(url_full):
    try:
        response = urllib.request.urlopen(url_full)
    except urllib.error.URLError as e:
        print("open page error:", e.reason)
    else:
        html = response.read()
        return BeautifulSoup(html, "lxml")


def get_courses_list():
    soup = get_url_soup("https://www.coursera.org/sitemap~www~courses.xml")
    if soup:
        return [url_full.text for url_full in soup.find_all('loc')]


def get_course_info(course_url):
    soup = get_url_soup(course_url)
    if soup:
        title = soup.find('meta', attrs={'property': 'og:title'}).attrs['content'].replace(" | Coursera", "")
        start_date = soup.find('div', attrs={'class': 'startdate rc-StartDateString caption-text'}).text
        language = soup.find('div', attrs={'class': 'rc-Language'}).text
        duration = len(soup.find_all('div', attrs={'class': 'week-heading body-2-text'}))
        rating = soup.find_all('div', attrs={'class': 'ratings-text headline-2-text'})
        if rating:
            rating = rating[0].contents[0].text
        else:
            rating = "No rating yet"
        return {'1_title': title, '2_date': start_date, '3_language': language, '4_weeks': duration, "5_rating": rating}


def output_courses_info_to_xlsx(courses_data):
    excel_filename = 'courses_info.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.title = "Coursera"
    courses_data_frame = pd.DataFrame(courses_data)
    for row in dataframe_to_rows(courses_data_frame, index=True, header=True):
        sheet.append(row)
    book.save(filename=excel_filename)

if __name__ == '__main__':
    courses_urls = get_courses_list()
    courses_info = [get_course_info(url) for url in courses_urls[:3]]  # save info for first 20 courses
    output_courses_info_to_xlsx(courses_info)