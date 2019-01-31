from bs4 import BeautifulSoup
from collections import namedtuple
from openpyxl import Workbook
from dateutil.parser import parse
import argparse
import os
import re
import requests
import sys
import xml.etree.ElementTree as ET
import json

course_info = namedtuple("course_info", [
    "title",
    "link",
    "lang",
    "start_date",
    "duration_weeks",
    "rating"
])


def fetch(url):
    return requests.get(url).text


def extract_courses_links_from_text(text):
    loc_elements = ET.fromstring(text).findall(
        ".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc")
    return map(lambda item: item.text, loc_elements)


def get_courses_list():
    xml_string = fetch("https://www.coursera.org/sitemap~www~courses.xml")
    return list(extract_courses_links_from_text(xml_string))


def get_graphql_data(soup):
    graphql_data_container = soup.select_one(
        'script[type="application/ld+json"]')
    graphql_data = json.loads(graphql_data_container.get_text())
    return graphql_data


def get_course_weeks_duration(soup, page_html, graphql_data):
    week_days = 7
    if not graphql_data:
        return
    try:
        start_date = parse(graphql_data["@graph"]
                           [2]["hasCourseInstance"]["startDate"])
        end_date = parse(graphql_data["@graph"][2]
                         ["hasCourseInstance"]["endDate"])
    except KeyError:
        return None

    days = (end_date - start_date).days
    return days / week_days


def get_course_start_date_string(soup, page_html, graphql_data):
    try:
        return graphql_data["@graph"][2]["hasCourseInstance"]["startDate"]
    except KeyError:
        return None


def get_course_lang(soup, page_html, graphql_data):
    if not graphql_data:
        return soup.select_one(".ProductGlance > :last-of-type h4").get_text()

    return graphql_data["@graph"][2]["inLanguage"]


def get_course_rating(soup, page_html, graphql_data):
    if graphql_data:
        return graphql_data["@graph"][1]["aggregateRating"]["ratingValue"]

    try:
        rating_el = soup.select_one(".AboutCourse [class*=StarRating] ~ span")
        return rating_el.get_text()
    except AttributeError:
        return None


def get_course_info(link, page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    graphql_data = get_graphql_data(soup)
    return course_info(
        title=soup.select_one(".BannerTitle h1").get_text(),
        link=link,
        lang=get_course_lang(soup, page_html, graphql_data),
        start_date=get_course_start_date_string(soup, page_html, graphql_data),
        duration_weeks=get_course_weeks_duration(
            soup, page_html, graphql_data),
        rating=get_course_rating(soup, page_html, graphql_data)
    )


def put_courses_to_workbook(workbook, courses):
    workbook = Workbook()
    worksheet = workbook.active

    courses = list(courses)

    fields = courses[0]._fields
    worksheet.append(fields)

    for line in courses:
        worksheet.append(
            list(map(lambda x: x or 'Unknown', line))
        )


def get_params():
    parser = argparse.ArgumentParser(
        description="""
            Coursera parser - it will parse information about the courses and put it to the file
        """
    )
    parser.add_argument("output", help="Path to result file")
    parser.add_argument("--limit", default=20, type=int, help="Limit courses")
    args = parser.parse_args()

    if args.limit < 1:
        parser.error("Limit should be > 1")
    if os.path.isfile(args.output):
        parser.error("File already exists")

    return args


if __name__ == "__main__":
    args = get_params()

    print("We're going collect information about {} course(s). Please wait".format(
        args.limit))

    courses_list = get_courses_list()[:args.limit]

    if not courses_list:
        sys.exit("No courses found")

    courses_info = map(
        lambda link: get_course_info(link, fetch(link)),
        courses_list
    )

    workbook = Workbook()
    put_courses_to_workbook(workbook, courses_info)
    try:
        workbook.save(args.output)
    except FileNotFoundError:
        sys.exit("Can't write to file")

    print("Check information in output file")
