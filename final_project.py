# importing libraries
import requests
import sys
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Sending a HTTP request to a URL
url = "https://eservices.minnstate.edu/registration/search/advancedSubmit.html?" \
      "campusid=305&searchrcid=0305&searchcampusid=305&yrtr=20205&subject=" \
      "ITEC&courseNumber=&courseId=&openValue=OPEN_PLUS_WAITLIST&delivery=" \
      "ALL&showAdvanced=&starttime=&endtime=&mntransfer=&credittype=" \
      "ALL&credits=&instructor=&keyword=&begindate=&site=&resultNumber=250"

# create Workbook object
book = Workbook()
sheet = book.active

# Create a data list to store the data
table_data = []


# takes in a string (ie. M, T, W, Th, Fri, Sat, Sun)
# and returns days of the week ( ie. Monday, ... , Sunday)
def convert_to_weekday(wkd):
    if wkd == 'M':
        return 'Monday'
    if wkd == 'T':
        return 'Tuesday'
    if wkd == 'W':
        return 'Wednesday'
    if wkd == 'Th':
        return 'Thursday'
    if wkd == 'Fri':
        return 'Friday'
    if wkd == 'Sat':
        return 'Saturday'
    if wkd == 'Sun':
        return 'Sunday'
    if wkd == 'n/a':
        return 'n/a'
    if wkd == 'M W':
        return 'Monday Wednesday'
    if wkd == 'T Th':
        return 'Tuesday Thursday'


# takes in a raw table data and analyzes all the rows
def grab_all_rows(course_table):
    # Get all the rows of table
    for tr in course_table.tbody.find_all("tr"):  # find all tr's from table's tbody
        local_data = []
        for td in tr.find_all("td"):  # find all td's in tr
            # remove any newlines and extra spaces from left and right
            local_data.append(td.text.replace('\n', ' ').strip())
        # Grab just part of the data that's relevant
        table_data.append(local_data[1:12])  # avoid empty data


# Takes in raw data and format it to fit into our needs
# and write the formatted data to a spreadsheet
def grab_table_content():
    for item in table_data:
        # convert list to fit the format : ID, Course number a dash section,
        # Course Title, Course Meets day of the week (ie. Wednesday instead of W),
        # Time course meets, credits for the course, and Instructor
        # print("raw item : ", item)
        temp = [item[0], item[2] + '-' + item[3], item[4],
                convert_to_weekday(item[6]), item[7], item[8], item[10]]
        print("formatted item :", temp)
        # append formatted data to the spread sheet
        sheet.append(temp)
    # save workbook
    book.save('courseTable.xlsx')
    book.close()


def main():
    # Make a GET request to fetch the raw HTML content
    try:
        response = requests.get(url)
        print(response.status_code)
    except requests.exceptions.ConnectionError:
        print("Connection Error")
        sys.exit(1)

    html_content = response.text

    # Parse the html content
    soup = BeautifulSoup(html_content, 'html.parser')
    print(soup.prettify()[250:325])  # print the parsed data of html

    # Traverse through the HTML tag, where the content resides
    # Get the table having the class resultsTable
    course_table = soup.find("table", id="resultsTable")
    print(course_table)
    # name the tile sheet
    sheet.title = "ITEC Course Schedule"
    # define column dimensions width accordingly
    sheet.column_dimensions['B'].width = 15
    sheet.column_dimensions['C'].width = 35
    sheet.column_dimensions['D'].width = 15
    sheet.column_dimensions['E'].width = 20
    sheet.column_dimensions['G'].width = 20

    # add column headings before adding data: these must be strings
    sheet.append(["ID", "Section", "Course Title", "Week date",
                  "Time", "Credits", "Instructor"])

    grab_all_rows(course_table)
    grab_table_content()


if __name__ == "__main__":
    main()
