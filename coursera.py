from lxml import etree
from bs4 import BeautifulSoup
from datetime import date
from openpyxl import Workbook

import requests
import random
import json


def get_courses_list():
    xml_page = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    root = etree.fromstring(xml_page.content)
    random_courses = [course[0].text for course in random.sample(set(root), 20)]
    return random_courses


def get_course_info(course_slug):
    full_path = 'https://www.coursera.org/learn/{}'.format(course_slug)
    html_doc = requests.get(full_path).content
    soup = BeautifulSoup(html_doc, 'html.parser')
    try:
        request_params = {"q" : "slug", "slug" : course_slug, "fields" : "upcomingSessionStartDate"}
        date_json = json.loads(requests.get("https://api.coursera.org/api/courses.v1/",
            params = request_params).text)['elements'][0]['upcomingSessionStartDate']
        course_start_date = date.fromtimestamp(date_json / 1000.0)
    except KeyError:
        request_params = {"q" : "slug", "slug" : course_slug, "fields" : "plannedLaunchDate"}
        date_json = json.loads(requests.get("https://api.coursera.org/api/courses.v1/",
            params = request_params).text)['elements'][0]['plannedLaunchDate']
        course_start_date = date_json
    try:
        course_rating = soup.find(class_ = 'ratings-text').text
    except AttributeError:
        course_rating = 'No rating'
    course_name = soup.find(class_ = 'course-name-text').text
    course_language = soup.find(class_ = 'language-info').text
    weeks_list = soup.find_all(class_ = 'week')
    course_weeks_amount = len(weeks_list) if weeks_list else 'No information'
    return course_name, course_rating, course_language, course_start_date, course_weeks_amount


def output_courses_info_to_xlsx(filepath):
    wb = Workbook()
    ws = wb.active
    arr = iter(['Name', 'Rating', 'Language', 'Start Date','Weeks Amount'])
    for cells in ws.iter_rows(max_row=1,max_col=5,min_row=1):
        for cell in cells:
            cell.value = next(arr)
    coursers_list = get_courses_list()
    for cells, course in zip(ws.iter_rows(max_row=21,max_col=5,min_row=2), coursers_list):
        course_slug = course.split('/')[-1]
        params = get_course_info(course_slug)
        for cell, param in zip(cells, params):
            cell.value = param
    wb.save(filepath)


if __name__ == '__main__':
    filepath = input('Enter file that will consist information about courses:\n')
    print('Start parsing of 20 any random Coursera courses. You need to wait a little bit...')
    output_courses_info_to_xlsx(filepath)
    print('Done. Data stored to {}'.format(filepath))
