#!/usr/bin/python
# coding: UTF-8

import json
import os
import xlsxwriter

name = "Notitle"
data = {}
workbook = xlsxwriter.Workbook('data/log.xlsx')
worksheet = workbook.add_worksheet()

def load_file(file_name):
	f = open('logs/' + file_name, 'r')
	line = f.readline()
	colum_num = 0
	name = "NoName"

	while line:
		if line[0] == "[":
			name = load_name(line)
		elif line[0] == "{":
			load_json(line, name, colum_num)
			colum_num += 4
		line = f.readline()
	f.close()

def load_name(line):
	return line[39:]

def load_json(line, name, colum_num):
	data = json.loads(line)
	write_data(data, name, colum_num)

def write_data(data, name, colum_num):
	bold = workbook.add_format({'bold': True})

	worksheet.write(0, colum_num, name, bold)
	worksheet.write(0, colum_num + 1, "Action", bold)
	worksheet.write(0, colum_num + 2, "Window", bold)

	worksheet.write(1, colum_num, "GyroUp")
	worksheet.write(1, colum_num + 1, data["GyroUp"]["action"])
	worksheet.write(1, colum_num + 2, data["GyroUp"]["window"])

	worksheet.write(2, colum_num, "GyroDown")
	worksheet.write(2, colum_num + 1, data["GyroDown"]["action"])
	worksheet.write(2, colum_num + 2, data["GyroDown"]["window"])

	worksheet.write(3, colum_num, "PinchIn")
	worksheet.write(3, colum_num + 1, data["PinchIn"]["action"])
	worksheet.write(3, colum_num + 2, data["PinchIn"]["window"])

	worksheet.write(4, colum_num, "PinchOut")
	worksheet.write(4, colum_num + 1, data["PinchOut"]["action"])
	worksheet.write(4, colum_num + 2, data["PinchOut"]["window"])

	worksheet.write(5, colum_num, "SwipeUp")
	worksheet.write(5, colum_num + 1, data["SwipeUp"]["action"])
	worksheet.write(5, colum_num + 2, data["SwipeUp"]["window"])

	worksheet.write(6, colum_num, "SwipeDown")
	worksheet.write(6, colum_num + 1, data["SwipeDown"]["action"])
	worksheet.write(6, colum_num + 2, data["SwipeDown"]["window"])

	worksheet.write(7, colum_num, "SwipeRight")
	worksheet.write(7, colum_num + 1, data["SwipeRight"]["action"])
	worksheet.write(7, colum_num + 2, data["SwipeRight"]["window"])

	worksheet.write(8, colum_num, "SwipeLeft")
	worksheet.write(8, colum_num + 1, data["SwipeLeft"]["action"])
	worksheet.write(8, colum_num + 2, data["SwipeLeft"]["window"])

	worksheet.write(9, colum_num, "SingleTap")
	worksheet.write(9, colum_num + 1, data["SingleTap"]["action"])
	worksheet.write(9, colum_num + 2, data["SingleTap"]["window"])

	worksheet.write(10, colum_num, "DoubleTap")
	worksheet.write(10, colum_num + 1, data["DoubleTap"]["action"])
	worksheet.write(10, colum_num + 2, data["DoubleTap"]["window"])

	worksheet.write(11, colum_num, "ButtonLeft")
	worksheet.write(11, colum_num + 1, data["ButtonLeft"]["action"])
	worksheet.write(11, colum_num + 2, data["ButtonLeft"]["window"])

	worksheet.write(12, colum_num, "ButtonRight")
	worksheet.write(12, colum_num + 1, data["ButtonRight"]["action"])
	worksheet.write(12, colum_num + 2, data["ButtonRight"]["window"])

files = os.listdir(os.path.abspath('logs'))

for file in files:
	load_file(file)

workbook.close()
