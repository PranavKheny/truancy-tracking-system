# Author: Lindsey Allen
# Date: March 3, 2025

# This script goes with simple_test_week_1, it assumes that you have the week 1 data and 
# includes test cases for increases in hours, decreasing hours, and a new student being on the list.
# Note that the total number of school hours has gone up for all students because another week of classes
# has happened, but this doesn't mean the number of absence hours has gone down.

import constructor

def main():
    # list data structure to hold all the students
    students = []
    # assume you know how many rows are on the pdf output; 4 lines would be 4 students
    lines = 5

    # initialize the number of students present on the pdf
    for _ in range (lines):
        students.append(constructor.Student.create_empty())

    # Fred has crossed the 25 hour threshold
    # Fred's hours have not changed since Week 1
    students[0].firstName = "Fred"
    students[0].lastName = "Jones"
    students[0].id = "111222"
    students[0].age = "4"
    students[0].grade = "KG"
    students[0].excused = "0"
    students[0].unexcused = "25"
    students[0].medical = "0"
    students[0].suspension = "0"
    students[0].schoolTotal = "125"
    students[0].attendingTotal = "95"
    students[0].absenceTotal = "25"


    # Daphne has crossed the 25 hour threshold and has medically excused hours
    # Medically excused hours are included in attendingTotal calculation but not absenceTotal
    # Daphne's unexcused hours have increased since Week 1
    students[1].firstName = "Daphne"
    students[1].lastName = "Blake"
    students[1].id = "333444"
    students[1].age = "6"
    students[1].grade = "01"
    students[1].excused = "0"
    students[1].unexcused = "30"
    students[1].medical = "5"
    students[1].suspension = "0"
    students[1].schoolTotal = "125"
    students[1].attendingTotal = "90"
    students[1].absenceTotal = "30"

    # Velma had more than 40 unexcused hours last week, this week she has less than 40 unexcused, 10 have been moved
    # to medically excused
    students[2].firstName = "Velma"
    students[2].lastName = "Dinkley"
    students[2].id = "555666"
    students[2].age = "7"
    students[2].grade = "02"
    students[2].excused = "2.11"
    students[2].unexcused = "33.9"
    students[2].medical = "10"
    students[2].suspension = "0"
    students[2].schoolTotal = "150"
    students[2].attendingTotal = "104.8"
    students[2].absenceTotal = "35.2"

    # Shaggy has crossed the 40 hour threshold and has some suspension hours
    # Suspension hours are like medical; they aren't included in absenceTotal
    # Shaggy's hours have not changed from last week.
    students[3].firstName = "Shaggy"
    students[3].lastName = "Rogers"
    students[3].id = "777888"
    students[3].age = "8"
    students[3].grade = "03"
    students[3].excused = "0"
    students[3].unexcused = "40"
    students[3].medical = "0"
    students[3].suspension = "8"
    students[3].schoolTotal = "150"
    students[3].attendingTotal = "102"
    students[3].absenceTotal = "40"

    # Scooby is new to the list
    # He has crossed the 25 hour threshold
    students[4].firstName = "Scooby"
    students[4].lastName = "Doo"
    students[4].id = "999000"
    students[4].age = "10"
    students[4].grade = "05"
    students[4].excused = "0"
    students[4].unexcused = "25"
    students[4].medical = "0"
    students[4].suspension = "0"
    students[4].schoolTotal = "150"
    students[4].attendingTotal = "95"
    students[4].absenceTotal = "25"

main()
