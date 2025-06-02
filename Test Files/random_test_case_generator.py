'''
	This file is used to randomly generate test data. When instantiated as a class the RandomTest object will keep track of what students have been randomly generated in the dataset and will vary the data according to parameters given by the user. It is capable of generating an arbitrary amount of students for an arbitrary number of weeks.
'''

import constructor
import random
import excel_writer
import openpyxl
from openpyxl import Workbook
import os
import datetime

class RandomTest:

	#source from which student names are randomly selected
	names = ["John", "Susan", "Marie", "Michael", "Kevin", "David", "Lauren", "Lord Sidious, God King of Mankind"]

	def __init__(self, count):
		self.fixed = []
		self.ids = []
		self.addStudents(count)
	
	#the method belows returns a student object with random data
	@staticmethod
	def createRandomStudent():
		return constructor.Student(random.randint(0,999999), RandomTest.names[random.randrange(0,len(RandomTest.names))], RandomTest.names[random.randrange(0,len(RandomTest.names))],
			random.randint(0,999), random.randint(0,999), random.randint(0,100), random.randint(0,100), random.randint(0,100), random.randint(0,100),
			random.randint(0,100), random.randint(0,100), random.randint(0,100))
	
	#adds "count" students to internal array ("self.fixed") ensuring unique student ids
	def addStudents(self, count):
		for i in range(count):
			randStudent = RandomTest.createRandomStudent()
			while randStudent.id in self.ids:
				randStudent = RandomTest.createRandomStudent()
			self.fixed.append(randStudent)
			self.ids.append(randStudent.id)
		
	#generate a new week of data varying attendance metrics by up to +/-"delta" with "change" existing students deleted and "change" new students added
	def newWeek(self, change=1, delta=10):
	
		#vary each students attendance metrics by +/-delta
		for i in range(len(self.fixed)):
			self.fixed[i].excused += random.randint(-delta, delta)
			self.fixed[i].unexcused += random.randint(-delta, delta)
			self.fixed[i].medical += random.randint(-delta, delta)
			self.fixed[i].suspension += random.randint(-delta, delta)
			self.fixed[i].schoolTotal += random.randint(-delta, delta)
			self.fixed[i].attendingTotal += random.randint(-delta, delta)
			self.fixed[i].absenceTotal += random.randint(-delta, delta)
		
		#remove change random students
		for i in range(change):
			self.fixed.pop(random.randrange(0, len(self.fixed)))
		
		#add change new students
		self.addStudents(change)
		
	#print method to display all student data visually
	def print(self):
		constructor.Student.printHeaders();
		for child in self.fixed:
			child.print()

#create new RandomTest with 5 students
newTest = RandomTest(5)
newTest.print()

#delete spreadsheet if exists
if os.path.isfile("goodGoing.xlsx"):
	os.remove("goodGoing.xlsx")

#create new spreadsheet
wb = Workbook()


#write first week to spreadsheet
excel_writer.writeWeek(wb, newTest.fixed, datetime.date.today() + datetime.timedelta(-7))
excel_writer.reorder_tabs(wb)

#run 7 more weeks
for i in range(7):
	print("###############################################################################")
	newTest.newWeek()
	newTest.print()
	excel_writer.writeWeek(wb, newTest.fixed, datetime.date.today() + datetime.timedelta(days=i*7))
	excel_writer.reorder_tabs(wb)


wb.save("goodGoing.xlsx")

