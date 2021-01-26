"""
#######################################################################################################################
###											Used for example of imports												###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import os
import re
import pandas as pd

# Meta Data
__author__ = 'David Mills'
__version__ = '0.0.1'
__email__ = 'david.mills@PSCconsulting.com'
__phone__ = '+44 7899 984158'
__status__ = 'Alpha'


def import_raw_load_estimates(pth_load_est, sheet_name='MASTER Based on SubstationLoad'):
	"""
		Function imports the raw load estimate into a DataFrame with no processing of the data
	:param str pth_load_est: Full path to file
	:param str sheet_name:  (optional) Name of worksheet in load estimate
	:return pd.DataFrame df_raw:
	"""

	# Read the raw excel worksheet
	df_raw = pd.read_excel(
		io=pth_load_est,			# Path to worksheet
		sheet_name=sheet_name,		# Name of worksheet to import
		skiprows=2,					# Skip first 2 rows since they do not contain anything useful
		header=0
	)

	# Remove any special characters from the column names (i.e. new line characters)
	df_raw.columns = df_raw.columns.str.replace('\n', '')

	return df_raw


def adjust_years(headers_list):
	"""
		Function will find the headers which contain the years associated with the forecast so that they can be
		duplicated for diversified and aggregate load
	:param list  headers_list:  List of all  headers
	:return list forecast_years:  List of headers now just for the forecast years
	"""

	# This is a Regex search string being compiled for matching, further details available here:
	# https://docs.python.org/3/library/re.html
	# Effectively:
	# r = Declares as a raw string
	#  (\d{4}) = 4 digits between 0-9 in a group
	#  \s* = 0 or more spaces
	#  [/] = / symbol
	r = re.compile(r'(\d{4})\s*[/]\s*(\d{4})')
	# The following extracts a list of all of the times the above is true
	forecast_years = filter(r.match, headers_list)

	return forecast_years


def get_local_file_path(file_name):
	"""
		Function returns the full path to a file which is stored in the same directory as this script
	:param str file_name:  Name of file
	:return str file_pth:  Full path to file assuming it exists in this folder
	"""
	local_dir = os.path.dirname(os.path.realpath(__file__))
	file_pth = os.path.join(local_dir, file_name)

	return file_pth


# The following statement is used by PyCharm to stop it being flagged as a potential error, should only be used when
# necessary
# noinspection PyClassHasNoInit
class Headers:
	"""
		Headers used as part of the DataFrame
	"""
	gsp = 'GSP'
	nrn = 'NRN'
	name = 'Name'
	voltage = 'Voltage Ratio'
	psse_1 = 'PSS/E Bus #1'

	# Columns to define attributes
	sub_gsp = 'Sub_GSP'
	sub_primary = 'Sub_Primary'

	# Header adjustments
	aggregate = 'aggregate'
