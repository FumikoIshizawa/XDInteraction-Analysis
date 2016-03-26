#!/usr/bin/python
# coding: UTF-8

import json
import os
import xlsxwriter

workbook = xlsxwriter.Workbook('data/log.xlsx')

gestures = [
	"GyroUp", 
	"GyroDown", 
	"PinchIn", 
	"PinchOut", 
	"SwipeUp", 
	"SwipeDown", 
	"SwipeLeft", 
	"SwipeRight", 
	"SingleTap", 
	"DoubleTap", 
	"ButtonLeft", 
	"ButtonRight"
]

def load_file(file_name):
	f = open('logs/' + file_name, 'r')
	line = f.readline()
	name = "NoName"
	column = 0
	start_row = 10
	custom_row = 0
	worksheets_num = {}
	pre_data = {}
	time = ""

	worksheet = workbook.add_worksheet()

	while line:
		if line[0] == "[":
			if line[39:44] == "Start" or line[39:44] == "Custo":
				message_data = line[39:]
				message_time = line[12:20]
				write_message(worksheet, message_data, custom_row, column, message_time)
				custom_row = 0 if custom_row > 9 else custom_row + 1
			else:
				name = load_name(line).strip()
				if name not in worksheets_num:
					worksheets_num[name] = start_row
					start_row += 14
					pre_data[name] = {
						"SwipeLeft":{"action":"Next Page","window":0},
						"GyroUp":{"action":"Scroll Up","window":0},
						"ButtonLeft":{"action":"No Gesture","window":0},
						"SwipeDown":{"action":"Text Small","window":0},
						"SingleTap":{"action":"No Gesture","window":0},
						"SwipeUp":{"action":"Text Big","window":0},
						"PinchIn":{"action":"Size Up","window":0},
						"SwipeRight":{"action":"Back Page","window":0},
						"PinchOut":{"action":"Size Down","window":0},
						"DoubleTap":{"action":"No Gesture","window":0},
						"GyroDown":{"action":"Scroll Down","window":0},
						"ButtonRight":{"action":"No Gesture","window":0}
					}
				time = line[12:20]
		elif line[0] == "{":
			data = json.loads(line)
			write_data(worksheet, data, name, worksheets_num[name], column, pre_data[name], time)
			column += 4
			custom_row = 0
			pre_data[name] = data
		line = f.readline()
	f.close()

def load_name(line):
	return line[39:]

def write_message(worksheet, data, row_num, colum_num, time):
	worksheet.write(row_num, colum_num, data)
	worksheet.write(row_num, colum_num + 1, time)

def write_data(worksheet, data, name, row_num, colum_num, pre_data, time):
	diff = workbook.add_format()
	diff.set_bg_color('yellow')

	worksheet.write(row_num, colum_num, name)
	worksheet.write(row_num, colum_num + 1, time)

	for gesture in gestures:
		worksheet.write(gestures.index(gesture) + 1 + row_num, colum_num, gesture)
		if pre_data[gesture]["action"] == data[gesture]["action"]:
			worksheet.write(gestures.index(gesture) + row_num + 1, colum_num + 1, data[gesture]["action"])
		else:
			worksheet.write(gestures.index(gesture) + row_num + 1, colum_num + 1, data[gesture]["action"], diff)

		if pre_data[gesture]["window"] == data[gesture]["window"]:
			worksheet.write(gestures.index(gesture) + row_num + 1, colum_num + 2, data[gesture]["window"])
		else:
			worksheet.write(gestures.index(gesture) + row_num + 1, colum_num + 2, data[gesture]["window"], diff)

files = os.listdir(os.path.abspath('logs'))

for file in files:
	load_file(file)

workbook.close()
