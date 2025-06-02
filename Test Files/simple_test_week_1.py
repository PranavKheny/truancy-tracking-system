# Author: Lindsey Allen
# Date: March 3, 2025

# This script contains test data for one week of truancy data.
# For testing purposes, it might make the most sense to treat this as the first output;
# act as if the Excel sheet we'd be merging on to is blank.

import constructor

def main():
    # list data structure to hold all the students
    students = []
    # assume you know how many rows are on the pdf output; 4 lines would be 4 students
    lines = 4

    # initialize the number of students present on the pdf
    for _ in range (lines):
        students.append(constructor.Student.create_empty())

    # Note: currently, I'm assuming that the parser will read everything in as a string, which 
    # should be fine with dynamic typing, but can be changed if we want.

    # Fred has crossed the 25 hour threshold
    students[0].firstName = "Fred"
    students[0].lastName = "Jones"
    students[0].id = "111222"
    students[0].age = "4"
    students[0].grade = "KG"
    students[0].excused = "0"
    students[0].unexcused = "25"
    students[0].medical = "0"
    students[0].suspension = "0"
    students[0].schoolTotal = "100"
    students[0].attendingTotal = "75"
    students[0].absenceTotal = "25"


    # Daphne has crossed the 25 hour threshold and has medically excused hours
    # Medically excused hours are included in attendingTotal calculation but not absenceTotal
    students[1].firstName = "Daphne"
    students[1].lastName = "Blake"
    students[1].id = "333444"
    students[1].age = "6"
    students[1].grade = "01"
    students[1].excused = "0"
    students[1].unexcused = "25"
    students[1].medical = "5"
    students[1].suspension = "0"
    students[1].schoolTotal = "100"
    students[1].attendingTotal = "70"
    students[1].absenceTotal = "25"

    # Velma has crossed the 40 hour threshold
    # Has some excused hours that are involved with the attendingTotal calculation
    # and absenceTotal calculation
    students[2].firstName = "Velma"
    students[2].lastName = "Dinkley"
    students[2].id = "555666"
    students[2].age = "7"
    students[2].grade = "02"
    students[2].excused = "2.11"
    students[2].unexcused = "43.9"
    students[2].medical = "0"
    students[2].suspension = "0"
    students[2].schoolTotal = "120"
    students[2].attendingTotal = "74.8"
    students[2].absenceTotal = "45.2"

    # Shaggy has crossed the 40 hour threshold and has some suspension hours
    # Suspension hours are like medical; they aren't included in absenceTotal
    students[3].firstName = "Shaggy"
    students[3].lastName = "Rogers"
    students[3].id = "777888"
    students[3].age = "8"
    students[3].grade = "03"
    students[3].excused = "0"
    students[3].unexcused = "40"
    students[3].medical = "0"
    students[3].suspension = "8"
    students[3].schoolTotal = "120"
    students[3].attendingTotal = "72"
    students[3].absenceTotal = "40"

main()
