#!/usr/bin/env python


# In[1]:

import pandas as pd
import numpy as np



def prepit(prodfile, invfile, mbfile, savename):
	prod = pd.read_excel(prodfile, index_col=None, na_values=['NA'])
	makebuy = pd.read_excel(mbfile, index_col=None, na_values=['null'])
	inv = pd.read_excel(invfile, index_col=None, na_values=['NA'])
	print('files in')
	prod = prod.groupby(['PARTNUM']).sum()
	prod = prod.reset_index()
	prep = pd.merge(inv, makebuy, on='PARTNUM', how='right')
	prep = pd.merge(prep, prod, on='PARTNUM', how='left')
	prep.fillna(0, inplace=True)
	print('merges done')
	prep['(IO) Available'] = (prep['QTYONHAND'] + prep['QTYONORDER'] - prep['QTYALLOCATED'])
	prep['Available'] = prep['(IO) Available'] - prep['QTYTOFULFILL']
	print('new columns added')
	prep.to_excel(savename, 'Sheet1')
	print('saved as', savename)
	print('prep done')

# prepit('product.xlsx','invoh.xlsx' ,'makebuy.xlsx' , 'pandomania.xlsx')