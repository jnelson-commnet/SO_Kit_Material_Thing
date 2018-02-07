
# coding: utf-8

# In[1]:

import pandas as pd
import numpy as np
import datetime
import os

# fall = pd.read_excel('forecastall.xlsx', index=None, na_value=['NA'])
# fio = pd.read_excel('forecastio.xlsx', index=None, na_value=['NA'])
# partid = pd.read_excel('partid.xlsx', index=None, na_value=['NA'])

def forecast_reports(forcall, forcio, partref):



	# In[2]:

	homey = os.path.abspath(os.path.dirname(__file__))
	tempForecastFile = os.path.join(homey, 'tempforecast.xlsx')




	# In[3]:

	forcall.drop(['QTYONHAND','QTYALLOCATED','QTYONORDER','QTYTOFULFILL','(IO) Available'], axis=1, inplace=True)


	# In[4]:

	forcall = pd.merge(forcall, partref, on='PARTNUM', how='left')


	# In[5]:

	allmfg = forcall[forcall['MOneeded'] != 0]


	# In[6]:

	allpurchase = forcall[forcall['Available']<0]


	# In[7]:

	allpurchase = allpurchase[['PARTNUM', 'PartDesc', 'Available', 'Make/Buy', 'AVGCOST']]


	# In[8]:

	allpurchase['Value'] = allpurchase['AVGCOST'] * allpurchase['Available']


	# In[9]:

	allpurchase = allpurchase[allpurchase['Make/Buy'] == 'Buy']
	allpurchase.sort_values(by='PARTNUM', axis=0, inplace=True)


	# In[10]:

	allmfg = allmfg[['PARTNUM', 'PartDesc', 'MOneeded', 'Make/Buy', 'AVGCOST']]


	# In[11]:

	allmfg['Value'] = allmfg['AVGCOST'] * allmfg['MOneeded']
	allmfg.sort_values(by='PARTNUM', axis=0, inplace=True)


	# In[12]:

	forcio.drop(['QTYONHAND', 'QTYALLOCATED', 'QTYONORDER', 'QTYTOFULFILL', 'Available'], axis=1, inplace=True)


	# In[13]:

	forcio = pd.merge(forcio, partref, on='PARTNUM', how='left')


	# In[14]:

	iomfg = forcio[forcio['MOneeded'] != 0]


	# In[15]:

	iopurchase = forcio[forcio['(IO) Available']<0]


	# In[16]:

	iopurchase = iopurchase[['PARTNUM', 'PartDesc', '(IO) Available', 'Make/Buy', 'AVGCOST']]


	# In[17]:

	iopurchase['Value'] = iopurchase['AVGCOST'] * iopurchase['(IO) Available']


	# In[18]:

	iopurchase = iopurchase[iopurchase['Make/Buy'] == 'Buy']
	iopurchase.sort_values(by='PARTNUM', axis=0, inplace=True)


	# In[19]:

	iomfg = iomfg[['PARTNUM', 'PartDesc', 'MOneeded', 'Make/Buy', 'AVGCOST']]


	# In[20]:

	iomfg['Value'] = iomfg['AVGCOST'] * iomfg['MOneeded']
	iomfg.sort_values(by='PARTNUM', axis=0, inplace=True)


	# In[28]:

	# today = datetime.datetime.now().strftime("%Y-%d-%B %I%M%p")


	# In[29]:

	# writer = pd.ExcelWriter(path='Z:\planning systems\Reports\Forecast All\Forecast All ' + today + '.xlsx', engine='xlsxwriter')

	# In[30]:

	# allpurchase.to_excel(writer, sheet_name='Inv_needs_for_Purchase', index=False)
	# allmfg.to_excel(writer, sheet_name='Manufacture_order_needs', index=False)


	# In[31]:

	# writer.save()


	# In[32]:

	# writer = pd.ExcelWriter(path='Z:\planning systems\Reports\Forecast Issued\Forecast IO ' + today + '.xlsx', engine='xlsxwriter')


	# In[33]:

	# iopurchase.to_excel(writer, sheet_name='IO_Inv_needs_for_Purchase', index=False)
	# iomfg.to_excel(writer, sheet_name='IO_Manufacture_order_needs', index=False)


	# In[34]:

	# writer.save()


	# In[ ]:

	writer = pd.ExcelWriter(path=tempForecastFile, engine='xlsxwriter')

	allpurchase.to_excel(writer, sheet_name='allpur')
	allmfg.to_excel(writer, sheet_name='allman')
	iopurchase.to_excel(writer, sheet_name='iopur')
	iomfg.to_excel(writer, sheet_name='ioman')

	writer.save()


# forecast_reports(fall, fio, partid)

