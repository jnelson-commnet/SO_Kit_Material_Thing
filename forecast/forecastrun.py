#!/usr/bin/env python

import pandas as pd
import numpy as np
import timeit
import os

def refresh_forecast_numbers():
	
	import reports
	'''
	print('There is a list of builds to add found at codes/buildlist.xlsx')
	print('There is a list of builds to ignore found at codes/ignorelist.xlsx')
	print('Would you like to include the "build" list in this forecast? (y or n)')
	check1 = input()
	print('Would you like to include the "ignore" list in this forecast? (y or n)')
	check2 = input()
	check1 = check1.lower()
	check2 = check2.lower()
	'''

	start = timeit.default_timer()
	import sys

	sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')

	print('running api...')

	import connecttest

	homey = os.path.abspath(os.path.dirname(__file__))

	myresults = connecttest.create_connection(homey, 'productquery.txt')
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, homey, 'product.xlsx')

	myresults = connecttest.create_connection(homey, 'partidquery.txt')
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, homey, 'partid.xlsx')

	myresults = connecttest.create_connection(homey, 'bomexploderquery.txt')
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, homey, 'bomexploder.xlsx')

	myresults = connecttest.create_connection(homey, 'invohquery.txt')
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, homey, 'invoh.xlsx')

	myresults = connecttest.create_connection(homey, 'makebuyquery.txt')
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, homey, 'makebuy.xlsx')


	print('running prep...')

	productFile = os.path.join(homey, 'product.xlsx')
	invohFile = os.path.join(homey, 'invoh.xlsx')
	makebuyFile = os.path.join(homey, 'makebuy.xlsx')
	pandomaniaFile = os.path.join(homey, 'pandomania.xlsx')
	bomExploderFile = os.path.join(homey, 'bomexploder.xlsx')
	forcastIoFile = os.path.join(homey, 'forecastio.xlsx')
	forcastAllFile = os.path.join(homey, 'forecastall.xlsx')
	PartIdFile = os.path.join(homey, 'partid.xlsx')

	first = timeit.default_timer()
	import prep
	prep.prepit(productFile, invohFile, makebuyFile, pandomaniaFile)





	second = timeit.default_timer()
	import bomexploder

	print('running bomexploder (io)...')

	bplo = pd.read_excel(bomExploderFile, index=None, na_value=['NA'])
	prepped = pd.read_excel(pandomaniaFile, index=None, na_value=['NA'])

	### trying to add or remove builds from the forecast.  Excel sheets have to have part numbers and quantities filled in correctly.

	'''
	edit = pd.DataFrame()
	if check1 == 'y':
		buildlist = pd.read_excel('Z:\planning systems\python versions\codes\\buildlist.xlsx', index=None, na_value=['NA'])
		edit = edit.copy().append(buildlist)
	if check2 == 'y':
		ignorelist = pd.read_excel('Z:\planning systems\python versions\codes\ignorelist.xlsx', index=None, na_value=['NA'])
		edit = edit.copy().append(ignorelist)

	### this section applies the edits to prepped before it runs through the forecast.

	print(len(edit))
	if len(edit) > 0:
		print('in build add and remove')
		edit = edit.groupby('PARTNUM').sum()
		edit = edit.reset_index()
		prepped = pd.merge(prepped.copy(), edit, on='PARTNUM', how='left')
		prepped.fillna(0, inplace=True)
		prepped['(IO) Available'] = prepped['(IO) Available'].copy() - prepped['buildcalc'].copy()
		prepped['Available'] = prepped['Available'].copy() - prepped['buildcalc'].copy()
		prepped.drop('buildcalc', axis=1, inplace=True)

	# writer = pd.ExcelWriter(path= 'Z:\planning systems\python versions\codes\postprepped.xlsx', engine='xlsxwriter')
	# prepped.to_excel(writer, sheet_name='sheet 1')
	# writer.save()
	'''


	ioavail = bomexploder.run_bomexploder(prepped, bplo, '(IO) Available')

	third = timeit.default_timer()

	print('running bomexploder (all)...')

	bplo = pd.read_excel(bomExploderFile, index=None, na_value=['NA'])
	prepped = pd.read_excel(pandomaniaFile, index=None, na_value=['NA'])

	### trying to add or remove builds from the forecast.  Excel sheets have to have part numbers and quantities filled in correctly.

	'''
	edit = pd.DataFrame()
	if check1 == 'y':
		buildlist = pd.read_excel('Z:\planning systems\python versions\codes\\buildlist.xlsx', index=None, na_value=['NA'])
		edit = edit.copy().append(buildlist)
	if check2 == 'y':
		ignorelist = pd.read_excel('Z:\planning systems\python versions\codes\ignorelist.xlsx', index=None, na_value=['NA'])
		edit = edit.copy().append(ignorelist)

	### this section applies the edits to prepped before it runs through the forecast.

	print(len(edit))
	if len(edit) > 0:
		print('in build add and remove')
		edit = edit.groupby('PARTNUM').sum()
		edit = edit.reset_index()
		prepped = pd.merge(prepped.copy(), edit, on='PARTNUM', how='left')
		prepped.fillna(0, inplace=True)
		prepped['(IO) Available'] = prepped['(IO) Available'].copy() - prepped['buildcalc'].copy()
		prepped['Available'] = prepped['Available'].copy() - prepped['buildcalc'].copy()
		prepped.drop('buildcalc', axis=1, inplace=True)

	# writer = pd.ExcelWriter(path= 'Z:\planning systems\python versions\codes\postprepped.xlsx', engine='xlsxwriter')
	# prepped.to_excel(writer, sheet_name='sheet 1')
	# writer.save()
	'''
	allavail = bomexploder.run_bomexploder(prepped, bplo, 'Available')

	fourth = timeit.default_timer()

	ioavail.to_excel(forcastIoFile, 'Sheet1')
	allavail.to_excel(forcastAllFile, 'Sheet1')

	partid = pd.read_excel(PartIdFile, index=None, na_value=['NA'])

	reports.forecast_reports(allavail, ioavail, partid)

	last = timeit.default_timer()






	print(start)
	print(first)
	print(second)
	print(third)
	print(fourth)
	print(last)
