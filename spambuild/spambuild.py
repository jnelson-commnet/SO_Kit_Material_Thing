import pandas as pd
import numpy as np
import os
import datetime
import savefun


# asking for input of part number

# print('Please enter part number as you see it in Fishbowl')
# print('eg: 016-450-10 r02 or 016-450-10 r03 will have different results')
# print('')
# prod = input()

# print('')
# print('Now enter the quantity you would like to check for build shortages')
# print('')
# qty = input()
# qty = int(qty)

homey = os.path.abspath(os.path.join(os.path.dirname( __file__ ), '..'))
forecastPath = os.path.join(homey, 'forecast')
spamBuildPath = os.path.join(homey, 'spambuild')

productFile = os.path.join(forecastPath, 'product.xlsx')
invohFile = os.path.join(forecastPath, 'invoh.xlsx')
makebuyFile = os.path.join(forecastPath, 'makebuy.xlsx')
pandomaniaFile = os.path.join(forecastPath, 'pandomania.xlsx')
bomExploderFile = os.path.join(forecastPath, 'bomexploder.xlsx')
forcastIoFile = os.path.join(forecastPath, 'forecastio.xlsx')
forcastAllFile = os.path.join(forecastPath, 'forecastall.xlsx')
PartIdFile = os.path.join(forecastPath, 'partid.xlsx')
tempForecastFile = os.path.join(forecastPath, 'tempforecast.xlsx')

holding_bplo = pd.read_excel(bomExploderFile, index=None, na_value=['NA'])
holding_mb = pd.read_excel(makebuyFile, index=None, na_value=['NA'])
holding_partid = pd.read_excel(PartIdFile, index=None, na_value=['NA'])
holding_inv = pd.read_excel(invohFile, index=None, na_value=['NA'])

holding_allpur = pd.read_excel(tempForecastFile, sheetname='allpur', index_col=None, na_values=['NA'])
holding_allman = pd.read_excel(tempForecastFile, sheetname='allman', index_col=None, na_values=['NA'])
holding_iopur = pd.read_excel(tempForecastFile, sheetname='iopur', index_col=None, na_values=['NA'])
holding_ioman = pd.read_excel(tempForecastFile, sheetname='ioman', index_col=None, na_values=['NA'])


def spam_all_the_builds(productlist):
	
	totalproductframe = pd.DataFrame()

	for everypart in productlist['PartNum']:
		prod = everypart
		qty = 1

		print(prod, 'started!')

		# # Commenting this out along with the sections further down that used this path.
		# filepathreports='Z:\planning systems\python versions\spambuild\\'


		bplo = holding_bplo.copy()
		mb = holding_mb.copy()
		partid = holding_partid.copy()
		inv = holding_inv.copy()

		print('Query sheets imported')


		bomexpl = savefun.basic_bom_explode(bplo, mb, prod)

		bomexpl = savefun.sum_bom(bomexpl)
		bomexpl = savefun.add_mb_to_bom(bomexpl, mb)

		print('BOM exploded')

		#############################
		# The Following is taken from the Bomexploder and I'm adding it to this as an afterthought.
		# The output of Buildable is very similar, so I want to merge them together.
		# Hopefully none of my messy variable naming conflicts....

		allpur = holding_allpur.copy()
		allman = holding_allman.copy()
		iopur = holding_iopur.copy()
		ioman = holding_ioman.copy()

		print('Forecast sheets imported')

		allman.rename(columns={'MOneeded':'All MOneeded'}, inplace=True)
		allpur.rename(columns={'Available':'All Available'}, inplace=True)
		ioman.rename(columns={'MOneeded':'(IO) MOneeded'}, inplace=True)



		tempall = allpur.copy().append(allman.copy())
		tempio = iopur.copy().append(ioman.copy())


		tempio.drop(['AVGCOST','Make/Buy','PartDesc','Value'], axis=1, inplace=True)


		combine = pd.merge(tempall.copy(), tempio.copy(), on='PARTNUM', how='left')



		combine.drop(['Value','PartDesc','Make/Buy','AVGCOST'], axis=1, inplace=True)




		bomtoforecast = pd.merge(bomexpl.copy(), combine.copy(), on='PARTNUM', how='left')




		bomtoforecast = pd.merge(bomtoforecast.copy(), partid.copy(), on='PARTNUM', how='left')



		bomtoforecast = pd.merge(bomtoforecast.copy(), inv.copy(), on='PARTNUM', how='left')




		bomtoforecast = bomtoforecast[['PARTNUM', 'PartDesc', 'QUANTITY', 'QTYONHAND', 'QTYALLOCATED', 'QTYONORDER', 'All Available', '(IO) Available', 'All MOneeded', '(IO) MOneeded', 'Make/Buy', 'AVGCOST']]

		bomtoforecast.rename(columns={'QUANTITY':'BOM qty'}, inplace=True)

		print('Forecast linked to BOM')

		# bomtoforecast is what the BOMExploder usually produces.
		# Later, we will merge it with the result of Buildable.
		# That's all I'm going to steal for now.
		#############################


		# flexybom is going to have the inventory change and reflect what is built, so we know what to keep running
		flexybom = pd.merge(bomexpl.copy(), inv[['PARTNUM','QTYONHAND']].copy(), on='PARTNUM', how='left')
		flexybom.fillna(0, inplace=True)

		# actualinv is not going to add FG back to inventory for each run, so we know total shortages
		actualinv = flexybom.copy()




		# setting the flexybom inventory at a negative qty for what we want to check
		flexybom.loc[flexybom['PARTNUM'] == prod, 'QTYONHAND'] = -qty


		# this isolates the negative "Make" parts, namely the one we just created
		flexymakes = flexybom[flexybom['Make/Buy'] == 'Make'].copy()
		flexymakes = flexymakes[flexymakes['QTYONHAND'] < 0].copy()
		flexylen = len(flexymakes)

		emptybom = list()

		tank = 0

		while tank < flexylen:
			# print(flexymakes)
			# print(flexylen, 'start')

			# getting the top item off flexymakes
			purt = flexymakes['PARTNUM'].iloc[tank]
			cutie = flexymakes['QTYONHAND'].iloc[tank]*(-1)

			# exploding just the top level BOM and multiplying it by build qty
			tempobom = savefun.bom_return(bplo, purt)
			tempomult = savefun.fg_to_multiplier(tempobom, cutie, purt)
			tempototal = savefun.bom_multiplier(tempobom, tempomult)

			# making raw good qtys negative to subtract from actualinv
			tempototal.rename(columns={'QUANTITY':'tempqty','RAW':'PARTNUM'}, inplace=True)
			tempoFG = tempototal[tempototal['FG'] == 10].copy()
			tempoRAW = tempototal[tempototal['FG'] != 10].copy()
			tempoRAW['tempqty'] = tempoRAW['tempqty']*(-1)

			# reconnecting total so that FG will be added to flexybom, and we won't rerun the same FG.
			tempototal = tempoFG.copy().append(tempoRAW.copy())

			# in case it's empty, we still need an FG to prevent re-run on flexybom.  emptybom records for later reference.
			if len(tempototal) == 0:
				tempototal = tempototal.append({'PARTNUM': purt, 'BOM': np.nan, 'FG': 10, 'tempqty': cutie}, ignore_index=True)
				emptybom.append(purt)

			# adding fg and subtracting raw from flexybom
			flexybom = pd.merge(flexybom.copy(), tempototal[['PARTNUM','tempqty']].copy(), on='PARTNUM', how='left')
			flexybom.fillna(0, inplace=True)
			flexybom['QTYONHAND'] = flexybom['QTYONHAND'].copy() + flexybom['tempqty'].copy()
			flexybom.drop('tempqty', axis=1, inplace=True)

			# subtracting raw from actualinv
			actualinv = pd.merge(actualinv.copy(), tempoRAW[['PARTNUM','tempqty']].copy(), on='PARTNUM', how='left')
			actualinv.fillna(0, inplace=True)
			actualinv['QTYONHAND'] = actualinv['QTYONHAND'].copy() + actualinv['tempqty'].copy()
			actualinv.drop('tempqty', axis=1, inplace=True)

			flexymakes = flexybom[flexybom['Make/Buy'] == 'Make'].copy()
			flexymakes = flexymakes[flexymakes['QTYONHAND'] < 0].copy()
			flexylen = len(flexymakes)

			# print(flexylen, 'end')

		# print(emptybom)
		emptay = pd.DataFrame()

		if len(emptybom) > 0:
			emptay = emptay.append(emptybom)

		# Adding a column that tells us what the total we can build is with that part.
		actualinv['Buildable'] = qty + ((actualinv['QTYONHAND']) / (actualinv['QUANTITY']))
		actualinv.rename(columns={'QUANTITY':'BOM qty','PARTNUM':'Part','QTYONHAND':'Post Build Inv'}, inplace=True)

		#########################
		# Now that we have actualinv (the result of Buildable), it's time to merge the bomexploder/forecast.
		# I'm afraid, so I'm making a copy.  It feels safer that way.
		reducebomex = bomtoforecast.copy()
		# Now I'm dropping the redundant columns to avoid merge complications.
		reducebomex.drop(['BOM qty','Make/Buy','AVGCOST'], axis=1, inplace=True)
		# Renaming the part number column to match (cus I wanna)
		reducebomex.rename(columns={'PARTNUM':'Part'}, inplace=True)

		theNewBuildable = pd.merge(actualinv.copy(), reducebomex.copy(), on='Part', how='left')
		theNewBuildable = theNewBuildable[['Part', 'PartDesc', 'BOM qty', 'QTYONHAND', 'QTYALLOCATED', 'QTYONORDER', 'All Available', '(IO) Available', 'All MOneeded', '(IO) MOneeded', 'Make/Buy', 'AVGCOST', 'Buildable', 'Post Build Inv']]

		# writer = pd.ExcelWriter(path=filepathreports + str(qty) + ' buildable ' + prod + '.xlsx', engine='xlsxwriter')

		# theNewBuildable.to_excel(writer, sheet_name='inv shortages', index=False)

		# emptay.to_excel(writer, sheet_name='missing boms', index=False)

		# writer.save()

		theNewBuildable['Originator'] = prod
		totalproductframe = totalproductframe.copy().append(theNewBuildable.copy())
		print(prod, 'done!')

	# # This shouldn't be necessary, just return the dataFrame.
	# writer = pd.ExcelWriter(path=filepathreports + 'Total Product.xlsx', engine='xlsxwriter')
	# totalproductframe.to_excel(writer, sheet_name='product_dump', index=False)
	# writer.save()

	return(totalproductframe)

