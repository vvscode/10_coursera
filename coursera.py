from bs4 import BeautifulSoup
from collections import namedtuple
from openpyxl import Workbook
import argparse
import os
import re
import requests
import sys
import xml.etree.ElementTree as ET

course_info = namedtuple("course_info", [
    "title",
    "link",
    "lang",
    "start_date",
    "duration_weeks",
    "rating"
])


def get_text_from_url(url):
    return requests.get(url).text


def extract_courses_links_from_text(text):
    loc_elements = ET.fromstring(text).findall(".//{http://www.sitemaps.org/schemas/sitemap/0.9}loc")
    return map(lambda item: item.text, loc_elements)


def get_courses_list():
    xml_string = get_text_from_url("https://www.coursera.org/sitemap~www~courses.xml")
    return list(extract_courses_links_from_text(xml_string))


def get_course_weeks_duration(page, page_html):
    week_begin_marker = "material.weeks."
    number_start = page_html.rfind(week_begin_marker) + len(week_begin_marker)
    number_end = page_html.find('"', number_start)
    number_string = page_html[number_start:number_end].strip()
    # numbers starts from 0
    return int(number_string) + 1


def get_course_start_date_string(page, page_html):
    return re.search(r'validFrom":"(.+?)"', page_html).group(1)


def get_course_lang(page, page_html):
    return page.select_one(".ProductGlance > :last-of-type h4").get_text()


def get_course_rating(page, page_html):
    try:
        return page.select_one(".AboutCourse [class*=StarRating] ~ span").get_text()
    except AttributeError:
        return "Unknown"


def get_course_info(link):
    page_html = get_text_from_url(link)
    page = BeautifulSoup(page_html, "html.parser")
    return course_info(
        title=page.select_one(".BannerTitle h1").get_text(),
        link=link,
        lang=get_course_lang(page, page_html),
        start_date=get_course_start_date_string(page, page_html),
        duration_weeks=get_course_weeks_duration(page, page_html),
        rating=get_course_rating(page, page_html)
    )


def output_courses_info_to_xlsx(filepath, courses):
    if not courses:
        return

    workbook = Workbook()
    worksheet = workbook.active

    fileds = courses[0]._fields
    worksheet.append(fileds)

    for line in courses:
        worksheet.append(line)

    workbook.save(filepath)


def get_params():
    parser = argparse.ArgumentParser(
        description="Coursera parser - it will parse information about the courses and put it to the file"
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

    print("We're going collect information about {} course(s). Please wait".format(args.limit))

    courses_list = get_courses_list()[:args.limit]

    if not courses_list:
        sys.exit("No courses found")

    courses_info = list(map(get_course_info, courses_list))

    try:
        output_courses_info_to_xlsx(args.output, courses_info)
    except FileNotFoundError:
        sys.exit("Can't write to file")

    print("Check information in output file")