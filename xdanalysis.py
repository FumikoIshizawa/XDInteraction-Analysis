#!/usr/bin/python
# coding: UTF-8

import json
import os
import xlsxwriter

workbook = xlsxwriter.Workbook('data/log.xlsx')
worksheet = workbook.add_worksheet()

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
	colum_num = 0
	pre_data = {
		"SwipeLeft":{"action":"Next Page","window":0},
		"GyroUp":{"action":"Scroll Up","window":0},
		"ButtonLeft":{"action":"No Gesture","window":0},
		"SwipeDown":{"action":"Text Small","window":0},
		"SingleTap":{"action":"No Gesture","window":0},
		"SwipeUp":{"action":"Text Big","window":0},
		"PinchIn":{"action":"Zoom In","window":0},
		"SwipeRight":{"action":"Back Page","window":0},
		"PinchOut":{"action":"Zoom Out","window":0},
		"DoubleTap":{"action":"No Gesture","window":0},
		"GyroDown":{"action":"Scroll Down","window":0},
		"ButtonRight":{"action":"No Gesture","window":0}
	}
	name = "NoName"

	while line:
		if line[0] == "[":
			name = load_name(line)
		elif line[0] == "{":
			data = json.loads(line)
			write_data(data, name, colum_num, pre_data)
			colum_num += 4
			pre_data = data
		line = f.readline()
	f.close()

def load_name(line):
	return line[39:]

def write_data(data, name, colum_num, pre_data):
	diff = workbook.add_format()
	diff.set_bg_color('yellow')

	worksheet.write(0, colum_num, name)
	worksheet.write(0, colum_num + 1, "Action")
	worksheet.write(0, colum_num + 2, "Window")

	for gesture in gestures:
		worksheet.write(gestures.index(gesture) + 1, colum_num, gesture)
		if pre_data[gesture]["action"] == data[gesture]["action"]:
			worksheet.write(gestures.index(gesture) + 1, colum_num + 1, data[gesture]["action"])
		else:
			worksheet.write(gestures.index(gesture) + 1, colum_num + 1, data[gesture]["action"], diff)

		if pre_data[gesture]["window"] == data[gesture]["window"]:
			worksheet.write(gestures.index(gesture) + 1, colum_num + 2, data[gesture]["window"])
		else:
			worksheet.write(gestures.index(gesture) + 1, colum_num + 2, data[gesture]["window"], diff)

files = os.listdir(os.path.abspath('logs'))

for file in files:
	load_file(file)

workbook.close()
