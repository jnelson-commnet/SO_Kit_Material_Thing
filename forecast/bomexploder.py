
import pandas as pd
import numpy as np
import datetime

# bplo = pd.read_excel('bomexploder.xlsx', index=None, na_value=['NA'])

# prepped = pd.read_excel('pandomania.xlsx', index=None, na_value=['NA'])

def run_bomexploder(fcast, bomsheet, shortcol):
	missbom = pd.DataFrame()
	fcast['MOneeded'] = 0
	while True == True:
		# get a list of shortages based on the All or IO columns
		tempshort = fcast[['PARTNUM', 'Make/Buy', shortcol]][fcast[shortcol]<0].copy()
		# limit the shortage list to Make items
		tempshort = tempshort[tempshort['Make/Buy'] == 'Make'].copy()
		# list out the existing FG lines of the BOM Exploder
		tempconvert = bomsheet[['RAW', 'BOM']][bomsheet['FG'] == 10].copy()
		# save a list of duplicate RAW parts in the FG lines
		bomdups = tempconvert[tempconvert.duplicated('RAW')].copy()
		# drop the duplicate FG lines
		tempconvert = tempconvert.drop_duplicates(subset='RAW').copy()
		# rename part column for easy merging
		tempconvert.rename(columns={'RAW':'PARTNUM'}, inplace=True)
		# merge list of BOMS onto list of shortages
		mshortconvert = pd.merge(tempshort.copy(), tempconvert.copy(), on='PARTNUM', how='left')
		# append rows missing BOM to missbom for reference
		missbom = missbom.append(mshortconvert[pd.isnull(mshortconvert['BOM']) == True])
		# drop any rows missing BOMs
		mshortconvert = mshortconvert.dropna()
		# if all of the Make shortages are missing BOMs then len(mshortconvert) == 0.
		# so save the list of duplicate FG's and missing BOM's and return the final forecast.
		if len(mshortconvert) == 0:
			print('done')
			today = datetime.datetime.now().strftime("%Y-%d-%B %I%M%p")
			# writer = pd.ExcelWriter(path='Z:\planning systems\Reports\BOM Checks\duplicates ' + shortcol + ' ' + today + '.xlsx', engine='xlsxwriter')
			# bomdups.to_excel(writer, sheet_name='duplicates')
			# writer.save()
			# writer = pd.ExcelWriter(path='Z:\planning systems\Reports\BOM Checks\missing ' + shortcol + ' bom ' + today + '.xlsx', engine='xlsxwriter')
			# missbom.to_excel(writer, sheet_name='no bom')
			# writer.save()
			return fcast
		else:
		    print('keep going!')
		# drop PARTNUM off shortages because all we need is the BOM to run
		mshortconvert.drop('PARTNUM', axis=1, inplace=True)
		# attach list of shortages by BOM to BOM Exploder so we know how many to run
		bommove = pd.merge(bomsheet.copy(), mshortconvert.copy(), on='BOM', how='left')
		# drop all N/A
		bommove.dropna(inplace=True)
		# multiply quantity per by shortage to get a number to adjust inventory by
		bommove['additive'] = bommove['QUANTITY'] * bommove[shortcol]
		# drop BOM, quantity per, and shortage column.  Should only leave Part, new additive number, and FG type.
		bommove.drop(['BOM', 'QUANTITY', shortcol], axis=1, inplace=True)
		# split out Finished Goods (10) and Raw Goods (20)
		fg = bommove[bommove['FG'] == 10].copy()
		raw = bommove[bommove['FG'] == 20].copy()
		# make copies?  I don't know why I didn't do this in the last step.
		fg = fg.copy()
		raw = raw.copy()
		# Switching the sign on FG's.
		fg['additive2'] = fg['additive'] * (-1)
		fg.drop(['FG', 'additive'], axis=1, inplace=True)
		raw.drop(['FG'], 1, inplace=True)
		raw = raw.groupby(['RAW']).sum()
		raw = raw.reset_index()
		fg = fg.groupby(['RAW']).sum()
		fg = fg.reset_index()
		# get these ready to merge onto the forecast
		raw.rename(columns={'RAW':'PARTNUM'}, inplace=True)
		fg.rename(columns={'RAW':'PARTNUM'}, inplace=True)
		# merge! merge damnit!
		fcast = pd.merge(fcast.copy(), raw.copy(), on='PARTNUM', how='left')
		fcast = pd.merge(fcast.copy(), fg.copy(), on='PARTNUM', how='left')
		fcast.fillna(0, inplace=True)
		# combine the values to a new master value
		fcast['newval'] = fcast[shortcol] + fcast['additive'] + fcast['additive2']
		# adding Manufacture orders to MOneeded (running tally of how many MO's need to be created)
		fcast['MOneededtemp'] = fcast['MOneeded'] + fcast['additive2']
		fcast.drop([shortcol, 'additive', 'additive2', 'MOneeded'], axis=1, inplace=True)
		fcast.rename(columns={'newval':shortcol}, inplace=True)
		fcast.rename(columns={'MOneededtemp':'MOneeded'}, inplace=True)






# run_bomexploder(prepped, bplo, '(IO) Available')


