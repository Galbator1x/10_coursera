import requests
import re
import io

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
    tree = html.fromstring(course_html)

    title = soup.find('div', class_='display-3-text')
    if title is None:
        return None

    language = 'missing'
    rows = len(tree.xpath('//*[@id=" "]/div/div[5]/table/tbody/tr'))
    for row in range(1, rows + 1):
        row_title = tree.xpath('//*[@id=" "]/div/div[5]/table/tbody/tr[{}]/td[1]/span/text()'.
                               format(str(row)))[0]
        if row_title == 'Language':
            language = tree.xpath('//*[@id=" "]/div/div[5]/table/tbody/tr[{}]/td[2]/span/span/text()'.
                                  format(str(row)))[0]
            break

    start_date = re.findall(r'"plannedLaunchDate":"([\w\d\s\.,-]+)"', course_html)
    if not start_date:
        start_date = re.findall(r'"startDate":"([\w\d\s\.,-]+)"', course_html)
    start_date = start_date[0] if start_date else 'missing'

    weeks_count = len(soup.find_all('div', class_='week'))

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
