

import os
import pandas as pd
import sys

homey = os.path.abspath(os.path.dirname(__file__))
forecastPath = os.path.join(homey, 'forecast')
spamBuildPath = os.path.join(homey, 'spambuild')

reportFile = os.path.join(homey, 'Result.xlsx')

sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')
import connecttest

sys.path.insert(0, forecastPath)
import forecastrun

sys.path.insert(0, spamBuildPath)
import spambuild

# forecastrun.refresh_forecast_numbers()

# myresults = connecttest.create_connection(homey, 'Shipping_Line_Query.txt')
# myexcel = connecttest.makeexcelsheet(myresults)
# connecttest.save_workbook(myexcel, homey, 'Shipping_Line_Query.xlsx')

shippingLineFile = os.path.join(homey, 'Shipping_Line_Query.xlsx')

shippingLines = pd.read_excel(shippingLineFile, index=None, na_value=['NA'])

shippingLines.drop_duplicates('PartNum', keep='first', inplace=True)

endFrame = spambuild.spam_all_the_builds(shippingLines.copy())

endFrame.drop(['BOM qty', 'All MOneeded', '(IO) MOneeded', 'AVGCOST', 'Buildable', 'Post Build Inv'], axis='columns', inplace=True)
endFrame.rename(columns={'All Available':'Forecasted Need', '(IO) Available':'Current Issued Shortage'}, inplace=True)
endFrame.drop_duplicates('Part', keep='first', inplace=True)
endFrame = endFrame[endFrame['Make/Buy'] == 'Buy'].copy()
endFrame.sort_values(by='Part', inplace=True)
endFrame.dropna(axis='index', how='any', subset=['Forecasted Need'], inplace=True)


writer = pd.ExcelWriter(reportFile, engine='xlsxwriter')
endFrame.to_excel(writer, sheet_name='product_dump', index=False)
writer.save()

