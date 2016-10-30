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
        date_json = json.loads(requests.get('https://api.coursera.org/api/courses.v1/?q=slug&slug=' +
        course_slug + '&fields=upcomingSessionStartDate').text)['elements'][0]['upcomingSessionStartDate']
        course_start_date = date.fromtimestamp(date_json / 1000.0)
    except:
        date_json = json.loads(requests.get('https://api.coursera.org/api/courses.v1/?q=slug&slug=' +
        course_slug + '&fields=plannedLaunchDate').text)['elements'][0]['plannedLaunchDate']
        course_start_date = date_json
    try:
        course_rating = soup.find(class_='ratings-text').text
    except:
        course_rating = 'No rating'
    course_name = soup.find(class_='course-name-text').text
    course_language = soup.find(class_='language-info').text
    weeks_list = soup.find_all(class_='week')
    course_weeks_amount = len(weeks_list) if weeks_list else 'No information'
    return course_name, course_rating, course_language, course_start_date, course_weeks_amount


def output_courses_info_to_xlsx(filepath):
    wb = Workbook()
    ws = wb.active
    for column, value in enumerate(['Name', 'Rating', 'Language', 'Start Date','Weeks Amount']):
        _ = ws.cell(column=column+1, row=1, value="{}".format(value))
    coursers_list = get_courses_list()
    for course_counter, course in enumerate(coursers_list):
        course_slug = course.split('/')[-1]
        course_values = get_course_info(course_slug)
        for column, value in enumerate(course_values):
            _ = ws.cell(column=column+1, row=course_counter+2, value="{}".format(value))
        print(' {0}/{1} courses parsed'.format(course_counter+1, len(coursers_list)), end='\r')
    wb.save(filepath)


if __name__ == '__main__':
    filepath = '{}.xlsx'.format(input('Enter file that will consist information about courses (without ".xlsx" part):\n'))
    print('Start parsing of 20 any random Coursera courses')
    output_courses_info_to_xlsx(filepath)
    print('Done. Data stored to {}'.format(filepath))
