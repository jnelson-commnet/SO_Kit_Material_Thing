#!/usr/bin/env python

import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
import datetime, calendar, openpyxl

### Returns a dataframe of the BOM that produces the FG of the input part
# bomdf: input bom exploder
# fgpart: input part number for FG of the BOM you want

def bom_return(bomdf, fgpart):
	fg = bomdf[bomdf['FG'] == 10].copy()
	# raw = bomdf[bomdf['FG'] == 20].copy()
	bomnum = fg['BOM'][fg['RAW'] == fgpart]
	if len(bomnum.index) > 0:
		bomnum = bomnum.iloc[0]
	else:
		bomnum = "nonononono"
	bomtemp = bomdf[bomdf['BOM'] == bomnum].copy()
	return bomtemp


### Returns multiplier necessary to run bom to match checkinv number
# singlebomdf: single FG bom exploder (comes from bom_return())
# checkinv: inventory number you want to produce
# fgpart: part number of the FG (in case there are multiple FG on a BOM, makes sure correct one is used)

def fg_to_multiplier(singlebomdf, checkinv, fgpart):
	fgqty = singlebomdf[singlebomdf['RAW'] == fgpart]
	fgqty = fgqty['QUANTITY'][fgqty['FG'] == 10].copy()
	if len(fgqty.index) > 0:
		fgqty = fgqty.iloc[0]
	else:
		fgqty = 1
	multiplier = checkinv/fgqty
	return multiplier

### Version of fg_to_multiplier before fgpart was included.  Keeping this here in case the patch ruins something.
# def fg_to_multiplier(singlebomdf, checkinv):
# 	fgqty = singlebomdf['QUANTITY'][singlebomdf['FG'] == 10].copy()
# 	if len(fgqty.index) > 0:
# 		fgqty = fgqty.iloc[0]
# 	else:
# 		fgqty = 1
# 	multiplier = checkinv/fgqty
# 	return multiplier


### Returns BOM with quantities multiplied by an input number
# bomdf: input bom exploder format
# multi: number to multiply by - not sure if floats will work, we'll see soon enough

def bom_multiplier(bomdf, multi):
	bomdf['mult'] = bomdf['QUANTITY'].copy() * multi
	bomdf.drop('QUANTITY', axis=1, inplace=True)
	bomdf.rename(columns={'mult':'QUANTITY'}, inplace=True)
	return bomdf.copy()


### Returns a df with P/N, Qty, and BOM for FG:
# bomfind: input bomexploder
# fgsearch: input part number you want for the FG
# output eg:
# 		|RAW			|QUANTITY	|	BOM
# 2598	|016-450-10 r02	|1.0		|	016-450-10 r02


def find_fg(bomfind, fgsearch):
	findit = bomfind[bomfind['FG'] == 10].copy()
	singlefgdf = findit[['RAW','QUANTITY','BOM']][findit['RAW'] == fgsearch].copy()
	return singlefgdf



### Returns a list of parts on a BOM:
# allboms: input full bomexploder
# makebuy: input makebuy sheet with AVGCOST included
# exactpart: input the part number you wish to run

# output eg:
# 	|BOM				|QUANTITY	|PARTNUM			|Make/Buy	
# 0	|016-450-10 r02		|1.0		|745-025-10 r01		|Buy		
# 1	|016-450-10 r02		|1.0		|590-785-10			|Buy		
# 2	|016-450-10 r02		|1.0		|590-538-10 r01		|Buy		

# NOTE: Uses find_fg()!! 

def basic_bom_explode(allboms, makebuy, exactpart):
	# getting a df started with just the PN and qty of 1
	runlist = find_fg(allboms, exactpart)
	# declare some variables for the loop
	buylist = pd.DataFrame()
	this = 0
	limiter = len(runlist.index)
	while this < limiter:
	    partcheck = runlist['RAW'].iloc[this]
	    partqty = runlist['QUANTITY'].iloc[this]
	    tempbom = bom_return(allboms, partcheck)
	    tempmult = fg_to_multiplier(tempbom, partqty, partcheck)
	    tempbom = bom_multiplier(tempbom.copy(), tempmult)
	    tempbom = pd.merge(tempbom.copy(), makebuy[['PARTNUM','Make/Buy']], left_on='RAW', right_on='PARTNUM')
	    tempbom = tempbom[tempbom['FG'] != 10].copy()
	    tempbuylist = tempbom[tempbom['Make/Buy'] == 'Buy'].copy()
	    temprunlist = tempbom[tempbom['Make/Buy'] == 'Make'].copy()
	    buylist = buylist.copy().append(tempbuylist.copy())
	    runlist = runlist.copy().append(temprunlist[['RAW','QUANTITY','BOM']])
	    # Added numbers to make increase the length of runlist
	    # print('limiter is', str(limiter), 'and this is', str(this))
	    limiter = len(runlist.index)
	    this+=1
	    

	splaybom = buylist.copy().append(runlist.copy())

	### I'm adjusting this part to dodge "Make" parts that are missing BOMs ... The commented area is the original, in case this is a mistake
	# splaybom.drop(['FG','Make/Buy','PARTNUM'], axis=1, inplace=True)
	# splaybom.rename(columns={'RAW':'PARTNUM'}, inplace=True)
	# splaybom = pd.merge(splaybom.copy(), makebuy[['PARTNUM','Make/Buy']], on='PARTNUM', how='left')
	# return splaybom
	### End of the original set.  The following part will ignore dataframes with a length of 0.

	if len(splaybom.index) > 0:
		splaybom.drop(['FG','Make/Buy','PARTNUM'], axis=1, inplace=True)
		splaybom.rename(columns={'RAW':'PARTNUM'}, inplace=True)
		splaybom = pd.merge(splaybom.copy(), makebuy[['PARTNUM','Make/Buy']], on='PARTNUM', how='left')
		return splaybom
	else:
		return pd.DataFrame()



### drops BOM and Make/Buy off the output of basic_bom_explode and sums QTYs:

def sum_bom(thisbom):
	thisbom = thisbom.groupby('PARTNUM').sum().copy()
	thisbom = thisbom.reset_index().copy()
	return thisbom



### adds Make/Buy and AVGCOST to list of PN:

def add_mb_to_bom(thisbom, makebuy):
	thisbom = pd.merge(thisbom.copy(), makebuy, on='PARTNUM', how='left')
	return thisbom



### takes a "Make" part number and inventory level to unbuild it into raw goods and add it to an inventory list:


def get_raw_goods_out(invsheet, allboms, partnumhere, partinvhere):
	# the usual, find bom, multiply it by inv
	tempbom = bom_return(allboms, partnumhere)
	tempmult = fg_to_multiplier(tempbom, partinvhere, partnumhere)
	tempbom = bom_multiplier(tempbom, tempmult)
	# seperate the FG from the tempbom
	fgtemp = tempbom[tempbom['FG'] == 10].copy()
	tempbom = tempbom[tempbom['FG'] != 10].copy()
	# make the FG negative
	fgtemp['QUANTITY'] = fgtemp['QUANTITY'].copy() * (-1)
	# re attach FG to tempbom
	tempbom = tempbom.copy().append(fgtemp.copy())
	# taking the number and adding it to invsheet
	tempbom.rename(columns={'RAW':'PARTNUM','QUANTITY':'calc'}, inplace=True)
	invsheet = pd.merge(invsheet.copy(), tempbom[['PARTNUM','calc']], on='PARTNUM', how='left')
	invsheet['calc'] = invsheet['calc'].fillna(0).copy()
	invsheet['QTYONHAND'] = invsheet['QTYONHAND'].fillna(0).copy()
	invsheet['QTYONHAND'] = invsheet['QTYONHAND'].copy() + invsheet['calc']
	invsheet.drop('calc', axis=1, inplace=True)
	return invsheet.copy()


### runs a list to put through get_raw_goods_out()

def make_list_add_inv(listtomake, bominvsheet, bomex):
	bominvsheet = bominvsheet.copy()
	for this in listtomake.index:
		partcheck = listtomake['PARTNUM'].ix[this]
		partqty = listtomake['QTYONHAND'].ix[this]
		bominvsheet = get_raw_goods_out(bominvsheet, bomex, partcheck, partqty)
	return bominvsheet


### produces a list of make parts with current inv from a bom with inv included (insert part number of FG so it can be removed):

def make_list_from_bom(bomwithinv, fgproduct):
	makeonly = bomwithinv[bomwithinv['Make/Buy'] == 'Make'].copy()
	makeonly = makeonly[makeonly['PARTNUM'] != fgproduct].copy()
	makeonly.drop(['QUANTITY','Make/Buy','AVGCOST'], axis=1, inplace=True)
	makeonly.dropna(inplace=True)
	return makeonly


### experiment with only doing one bom level at a time:

### takes a "Make" part number and inventory level to unbuild it into raw goods and add it to an inventory list:
	# Also adds more to the runlist
def get_raw_goods_out_add_to_make_list(bominvsheet, allbplo, makebuydf, pnhere, pinvhere):
	# the usual, find bom, multiply it by inv
	temporarybom = bom_return(allbplo, pnhere)
	multiplybythis = fg_to_multiplier(temporarybom, pinvhere, pnhere)
	temporarybom = bom_multiplier(temporarybom, multiplybythis)
	# seperate the FG from the temporarybom
	fingoodtemp = temporarybom[temporarybom['FG'] == 10].copy()
	temporarybom = temporarybom[temporarybom['FG'] != 10].copy()
	# save this for later to add make parts to run through list
	makecheck = temporarybom.copy()
	# make the FG negative
	fingoodtemp['QUANTITY'] = fingoodtemp['QUANTITY'].copy() * (-1)
	# re attach FG to temporarybom
	temporarybom = temporarybom.copy().append(fingoodtemp.copy())
	# taking the number and adding it to bominvsheet
	temporarybom.rename(columns={'RAW':'PARTNUM','QUANTITY':'calc'}, inplace=True)
	bominvsheet = pd.merge(bominvsheet.copy(), temporarybom[['PARTNUM','calc']], on='PARTNUM', how='left')
	bominvsheet['calc'] = bominvsheet['calc'].fillna(0).copy()
	bominvsheet['QTYONHAND'] = bominvsheet['QTYONHAND'].fillna(0).copy()
	bominvsheet['QTYONHAND'] = bominvsheet['QTYONHAND'].copy() + bominvsheet['calc']
	bominvsheet.drop('calc', axis=1, inplace=True)
	# want to do the make list now
	makecheck = pd.merge(makecheck.copy(), makebuydf[['PARTNUM','Make/Buy']].copy(), left_on='RAW', right_on='PARTNUM')
	makecheck = makecheck[makecheck['Make/Buy'] == 'Make'].copy()
	makecheck.drop(['RAW','BOM','FG','Make/Buy'], axis=1, inplace=True)
	return (bominvsheet.copy(), makecheck.copy())


###

# part_to_product_reference uses the "prepped" sheet and to find any part that is short and a "Make" part.
# Then it explodes the BOM for each of those parts and appends each of them to the same dataframe.

# Columns produced are:
# BOM : Each part that was used from the prep sheet
# Make/Buy : The Make/Buy state of the PARTNUM column
# PARTNUM : The part numbers for the raw goods on each of the BOMs
# QUANTITY : The amount of the PARTNUM used on said BOM


def part_to_product_reference(bomsheethere, makebuysheethere, prepsheethere, availcol):
	lowpostprep = prepsheethere[prepsheethere[availcol] < 0].copy()

	lowpostprep = lowpostprep[lowpostprep['Make/Buy'] == 'Make'].copy()

	reference = pd.DataFrame()
	for each in lowpostprep['PARTNUM']:
		bomex = basic_bom_explode(bomsheethere, makebuysheethere, each)
		bomex['BOM'] = each
		reference = reference.copy().append(bomex.copy())
	return reference.copy()
