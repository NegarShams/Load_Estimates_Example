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
import pandas as pd

# Unique imports
import common_functions as common

# GLOBAL constants
# Target filename to use
FILE_NAME_INPUT = '2019-20 SHEPD Load Estimates - v6.xlsx'
FILE_NAME_OUTPUT = 'Processed Load Estimates.xlsx'
FILE_PTH_INPUT = common.get_local_file_path(file_name=FILE_NAME_INPUT)
FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)


# Functions
def assign_gsp(df_raw):
	"""
		Function to assign a GSP to each Primary substation included within the dataset
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df:
	"""
	# Using Forward Fill to populate the GSP associated with each Primary down.
	# TODO: Potential risk here if the GSP name box is empty then an error will occur
	# Only forward fills where the GSP row has been identified by the sub_gsp column
	df_raw[common.Headers.gsp] = df_raw[common.Headers.gsp].where(df_raw[common.Headers.sub_gsp] == True).ffill()

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
	idx = (
			df_raw[common.Headers.name].isna() &		# Confirm that no Primary substation name entry exists
			~df_raw[common.Headers.gsp].isna() &		# Confirm that a GSP name entry exists
			~df_raw[common.Headers.voltage].isna()		# Confirm that a voltage ratio for the substation exists
			# TODO: Checking the voltage ratio may not be reliable
	)
	# The following line sets those rows under the column sub_gsp to True
	df_raw.loc[idx, common.Headers.sub_gsp] = True

	# Find the situations where there is an entry in the Name but no entry in the GSP column but that there is an entry
	# in the NRN column as this determines that it is a Primary substation and not other data.
	# TODO: In case an NRN number is missing an alternative check is that there is a PSSE busbar number
	#  TODO: in column PSSE_bus 1 but this will require some further checking for other conditions
	idx = (
			~df_raw[common.Headers.name].isna() & 	# Confirm that an entry is included for Primary substation name
			df_raw[common.Headers.gsp].isna() & 	# Confirm that no entry exists for the GSP substation name
			~df_raw[common.Headers.nrn].isna()		# Confirm that an entry is included for the NRN number
	)
	# The following line sets those rows under the column sub_gsp to True
	df_raw.loc[idx, common.Headers.sub_primary] = True

	return df_raw


def extract_aggregate_demand(df_raw):
	"""
		Extract the aggregate demand from the diversified demand for each GSP
	:param pd.DataFrame df_raw:
	:return pd.DataFrame df_raw:
	"""
	# Find the years being considered for the forecast
	forecast_years = common.adjust_years(headers_list=list(df_raw.columns))

	# Adjust the list to include a leading string value to identify this as a certain type of forecast (i.e aggregate)
	# TODO: Could make use of MultiIndex Pandas DataFrame Columns instead which would allow for more efficient filtering
	adjusted_list = ['{}_{}'.format(common.Headers.aggregate, x) for x in forecast_years]
	# Add columns to DataFrame with no files by adding in empty
	df_raw = pd.concat([df_raw, pd.DataFrame(columns=adjusted_list)], sort=False)

	# For columns which have been identified as GSP extract the aggregate demand from the row below and add to the GSP
	# row under the new sections for aggregate demand
	idx_gsp = df_raw[df_raw[common.Headers.sub_gsp] == True].index
	# loop through these index values for each GSP and get the aggregate values from the row below
	for idx in idx_gsp:
		df_raw.loc[idx, adjusted_list] = df_raw.loc[idx + 1, forecast_years].values

	return df_raw


def remove_unnecessary_rows(df_raw):
	"""
		Function removes all of the rows which do not correspond to the usable data for GSP or Primary substations
	:param pd.DataFrame df_raw: Input DataFrame to be processed
	:return pd.DataFrame df_out:  Output DataFrame after processing
	"""

	# Get list of all rows which are determined as either a GSP or Primary substation
	idx = (
		(df_raw[common.Headers.sub_gsp] == True) |			# Confirm if the row is a GSP
		(df_raw[common.Headers.sub_primary] == True)		# Confirm if the row is a Primary
	)

	# Return subset of the original DataFrame including only these rows
	df_out = df_raw[idx]

	return df_out


if __name__ == '__main__':
	# Import raw DF
	df = common.import_raw_load_estimates(pth_load_est=FILE_PTH_INPUT)
	# Identify whether a GSP or Primary substation for each row
	df = determine_gsp_primary_flag(df_raw=df)

	# Extract aggregate demand for each GSP
	df = extract_aggregate_demand(df_raw=df)

	df = assign_gsp(df_raw=df)

	df = remove_unnecessary_rows(df_raw=df)

	# Export processed DataFrame
	df.to_excel(FILE_PTH_OUTPUT)
