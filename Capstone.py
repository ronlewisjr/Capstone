import pandas as pd
import numpy as np

from pandas import datetime
from matplotlib import pyplot
from pandas.tools.plotting import autocorrelation_plot

import os
import cx_Oracle

from sklearn.model_selection import train_test_split
from sklearn.model_selection import KFold
from statsmodels.tsa.arima_model import ARIMA
from sklearn.metrics import mean_squared_error

###
### The function below converts the date time stamp from a file into a date without a time associated
### with it
###

def parser(x):
	return datetime.strptime(x, '%m/%d/%Y')

df = pd.read_csv('data/item_history.csv',header=0, parse_dates=[5], index_col=[5] )
df["REPLACEMENT_ITEM"].fillna('N', inplace=True)
df["QTY"]= round(df["QTY"]/df["QTY_PER_SELL_UOM"],4)
df = df.drop(["QTY_PER_SELL_UOM", "E3_YEAR", "E3_4WK_PERIOD"], axis=1)

#print (df.head(5))

#print (df)


for k1, group in df.groupby(["DIST_NO","ITEM_NO"]):
	print (k1)
	#print (group)
	dfs = pd.DataFrame(group)
	dfs = dfs.drop(["ITEM_NO","DIST_NO","HITS","REPLACEMENT_ITEM"], axis=1)
	print (dfs)
	
	dfs.plot()
	pyplot.show()

	autocorrelation_plot(dfs)
	pyplot.show()

	model = ARIMA(dfs, order=(5,1,0))
	model_fit = model.fit(disp=0)
	print(model_fit.summary())
	# plot residual errors
	residuals = pd.DataFrame(model_fit.resid)
	residuals.plot()
	pyplot.show()
	residuals.plot(kind='kde')
	pyplot.show()
	print(residuals.describe())


