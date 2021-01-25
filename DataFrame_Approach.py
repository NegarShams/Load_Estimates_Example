"""
#######################################################################################################################
###											Example DataFrame Processing											###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import os
import pandas as pd

# Unique imports
import common_functions as common

# GLOBAL constants
# Target filename to use
FILE_NAME = '2019-20 SHEPD Load Estimates - v6.xlsx'
FILE_PTH = common.get_local_file_path(file_name=FILE_NAME)

# Functions
def assign_gsp(df_raw):
	"""
		Function to assign a GSP to each Primary substation included within the dataset
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df:
	"""
	# Using Forward Fill to populate the GSP associated with each Primary down.
	# TODO: Potential risk here if the GSP name box is empty then an error will occur
	# inplace=True means that the DataFrame is updated rather than a new DataFrame being created
	df_raw[common.Headers.gsp].ffill(inplace=True)

	return df_raw


def determine_gsp_primary_flag(df_raw):
	"""
		Determines whether a row contains GSP or Primary substation data
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df:
	"""

	# If no entry in Name but an entry in GSP then assume GSP substation
	# TODO: There may be situations where this rule is not true
	# The following line determines all the rows for which this is True
	idx = df_raw[common.Headers.name].isna() & ~df_raw[common.Headers.gsp].isna()
	# The following line sets those rows under the column sub_gsp to True
	df_raw.loc[idx, common.Headers.sub_gsp] = True

	# Find the situations where there is an entry in the Name but no entry in the GSP column
	idx = ~df_raw[common.Headers.name].isna() & df_raw[common.Headers.gsp].isna()
	# The following line sets those rows under the column sub_gsp to True
	df_raw.loc[idx, common.Headers.sub_primary] = True

	return df

def extract_aggregate_demand(df_raw):
	"""
		Extract the aggregate demand from the diversified demand for each GSP
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df:
	"""
	# Find the years being considered for the forecast
	forecast_years = common.adjust_years(headers_list=list(df_raw.columns))

	# Adjust the list to include a leading string value to identify this as a certain type of forecast (i.e aggregate)
	# TODO: Could make use of MultiIndex Pandas DataFrame Columns instead which would allow for more efficient filtering
	adjusted_list = ['{}_{}'.format(common.Headers.aggregate, x) for x in forecast_years]
	# Add columns to DataFrame with no files by adding in empty
	df = pd.concat([df_raw, pd.DataFrame(columns=adjusted_list)])

	# For columns which have been identified as GSP extract the aggregate demand from the row below and add to the GSP
	# row under the new sections for aggregate demand
	idx_gsp = df[df[common.Headers.sub_gsp]==True].index
	# loop through these index values
	raise SyntaxError('Follow loop is not quite correct')
	for idx in idx_gsp:
		df.loc[idx, adjusted_list] = df.loc[idx+1, forecast_years]

	return df_raw


if __name__ == '__main__':
	# Import raw DF
	df = common.import_raw_load_estimates(pth_load_est=FILE_PTH)
	# Identify whether a GSP or Primary substation for each row
	df = determine_gsp_primary_flag(df_raw=df)

	# Extract aggregate demand for each GSP
	df = extract_aggregate_demand(df_raw=df)

	df = assign_gsp(df_raw=df)
