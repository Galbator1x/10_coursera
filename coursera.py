import requests
import re
import io
import json

from lxml import etree, html
from bs4 import BeautifulSoup
from openpyxl import Workbook


COURSES_XML_URL = 'https://www.coursera.org/sitemap~www~courses.xml'
QUANTITY_COURSES_TO_OUTPUT = 20


def get_courses_list():
    courses_xml = requests.get(COURSES_XML_URL).content
    tree = etree.parse(io.BytesIO(courses_xml))
    root = tree.getroot()
    courses_urls = [url[0].text for url in root]

    courses_list = []
    for url in courses_urls[:QUANTITY_COURSES_TO_OUTPUT]:
        course = get_course_info(url)
        if course is not None:
            courses_list.append(course)
    return courses_list


def get_course_info(course_slug):
    course_html = requests.get(course_slug).text
    soup = BeautifulSoup(course_html, 'html.parser')

    title = soup.find('div', class_='display-3-text')
    if title is None:
        return None

    language = 'missing'
    table = soup.find('table', class_='basic-info-table')
    td_list = table.find_all('td')
    for td_id, td in enumerate(td_list):
        try:
            if td.find('span').text == 'Language':
                language = re.findall(r'[\w\d\s\(\)]+',
                                      td_list[td_id + 1].find('span').text)[0]
                break
        except AttributeError:
            pass

    try:
        data_from_script = soup.select('script[type="application/ld+json"]')[0].text
        data_json = json.loads(data_from_script)
        start_date = data_json['hasCourseInstance'][0]['startDate']
    except (KeyError, IndexError):
        start_date = 'missing'

    weeks_count = len(soup.find_all('div', class_='week'))
    weeks_count = 'missing' if weeks_count == 0 else weeks_count

    rating = soup.find('div', class_='ratings-text')
    rating = re.findall(r'\d\.\d', rating.text)[0] if rating is not None else None

    return title.text, language, start_date, weeks_count, rating


def save_courses_info_to_xlsx(courses_list, filepath):
    wb = Workbook()
    worksheet = wb.active
    worksheet.append(['Title', 'Language', 'Start date', 'Weeks count', 'Rating'])
    [worksheet.append(course) for course in courses_list]
    wb.save(filepath)


if __name__ == '__main__':
    filepath = 'courses.xlsx'
    try:
        save_courses_info_to_xlsx(get_courses_list(), filepath)
    except requests.exceptions.ConnectionError:
        print('Connection aborted, try later.')
