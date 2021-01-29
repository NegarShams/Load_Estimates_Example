"""
#######################################################################################################################
###											Example DataFrame Comparison											###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

# Generic Imports
import pandas as pd
import numpy as np
import unittest

# Unique imports
import common_functions as common

# General Constants
FILE_NAME_OUTPUT = 'Example Comparison.xlsx'
FILE_NAME_OUTPUT2 = 'Example Comparison2.xlsx'
FILE_PTH_OUTPUT = common.get_local_file_path(file_name=FILE_NAME_OUTPUT)
FILE_PTH_OUTPUT2 = common.get_local_file_path(file_name=FILE_NAME_OUTPUT2)

# Engine to use when writing excel workbooks (XlsxWriter needed for formatting of tabs)
excel_engine = 'xlsxwriter'


# Functions
def produce_dataframe(dimensions, row_num):
	"""
		Produces a DataFrame where the row number is doubled
	:param int dimensions: Number of rows and columns
	:param int row_num: Row number that will be doubled
	:return pd.DataFrame df_raw:
	"""
	# Confirm that number of rows matches with dimensions and if not raise error message
	if row_num > dimensions:
		raise ValueError('Target row {} greater than dimensions {}'.format(row_num, dimensions))

	# Produces a list of lists of integers as generic data population
	generic_data = [range(0, dimensions) for _ in range(0, dimensions)]
	# Edit data to double a specific row
	generic_data[row_num] = [x * 2 for x in generic_data[row_num]]

	# Produces some column headers
	col_headers = ['Column {}'.format(x) for x in range(0, dimensions)]

	# Produce DataFrame
	df_raw = pd.DataFrame(data=generic_data, columns=col_headers)

	return df_raw


def highlight_diff(full_dataset, cells_to_highlight, color='yellow'):
	"""
		Small function to deal with highlighting the cells in the excel output that have changed from the original
		data set
	:param pd.DataFrame full_dataset:  Full dataframe
	:param pd.DataFrame cells_to_highlight:  DataFrame containing entries in only the cells which need highlighting
	:param str color:  Colour to highlight the cells
	"""
	# Set the attribute based on the colour input
	attr = 'background-color: {}'.format(color)
	print(attr)
	# Is difference
	is_diff = full_dataset == cells_to_highlight
	# Return a DataFrame with the style matching
	return pd.DataFrame(
		np.where(is_diff, attr, ''),
		index=full_dataset.index, columns=full_dataset.columns)


def compare_dataframes(df1, df2):
	"""
		Function compares two DataFrames of the same size and returns a dataframe of the same dimensions but with only
		the differences shown.  Keeps values in df2
	:param pd.DataFrame df1:  DataFrame 1
	:param pd.DataFrame df2:  DataFrame 2
	:return (pd.DataFrame, pd.DataFrame.Styled) (df_diff, df2_styled):  Different values between DataFrames and
																		original dataframe with changes highlighted
	"""

	# Confirm that DataFrames are the same dimensions otherwise raise error
	if df1.shape != df2.shape:
		raise ValueError('The two DataFrames provided are not the same dimensions ({} != {})'.format(df1.shape, df2.shape))

	# Keeps the values in df2 which are different to df1, all other values are set to nan
	local_df_diff = df2.where(df2 != df1)

	# Highlight the cells in df2 which have changed
	df2_styled = df2.style.apply(highlight_diff, cells_to_highlight=local_df_diff, axis=None)

	return local_df_diff, df2_styled


def write_dataframe(workbook, df, sheet_name, tab_color=None):
	"""
		Function deals with writing a DataFrame to new worksheet in excel whilst also formatting the worksheet tab
		color
	:param pd.ExcelWriter workbook:  Handle to workbook writer to use
	:param pd.DataFrame df:  Data to be written (if is a df.Styled) then will include highlighting and colour)
	:param str sheet_name:  Name to give worksheet
	:param str tab_color:  Color for tab
	:return: None
	"""
	# Create a new worksheet
	wksh = workbook.book.add_worksheet(name=sheet_name)

	# Have to add worksheet to Pandas list of worksheets
	# (https://stackoverflow.com/questions/32957441/putting-many-python-pandas-dataframes-to-one-excel-worksheet)
	workbook.sheets[sheet_name] = wksh

	# Only set tab_color if not None
	if tab_color:
		wksh.set_tab_color(tab_color)

	# Write DataFrame to excel worksheet
	df.to_excel(workbook, sheet_name=sheet_name)
	return None


# Some unit tests
class UnitTestExample(unittest.TestCase):

	def setUp(self):
		""" Some setup for the unittest, for example where the same variable input is used across multiple tests"""
		self.df_dimensions = 5

	def testDataFrameProduction(self):
		""" Confirms that a DataFrame is produced of the correct dimensions"""
		test_df = produce_dataframe(dimensions=self.df_dimensions, row_num=3)

		# Confirm the returned type is a DataFrame
		self.assertEqual(type(test_df), pd.DataFrame)
		# Confirm the dimensions of the DataFrame are as expected
		self.assertEqual(test_df.shape, (self.df_dimensions, self.df_dimensions))

	def testDataFrameProduction_Fails(self):
		""" Confirms that a DataFrame production fails if the input row_num is larger than the dimensions"""
		with self.assertRaises(ValueError):
			_ = produce_dataframe(dimensions=self.df_dimensions, row_num=self.df_dimensions + 2)

	def testDataFramesDifferent(self):
		""" Confirms that two DataFrames with different ids look different"""

		test_df1 = produce_dataframe(dimensions=self.df_dimensions, row_num=2)
		test_df2 = produce_dataframe(dimensions=self.df_dimensions, row_num=4)

		self.assertFalse(test_df1.equals(test_df2))

		# Retrieve differences
		test_df_diff, test_df2 = compare_dataframes(df1=test_df1, df2=test_df2)

		# Confirm the same shape as df1
		self.assertEqual(test_df1.shape, test_df_diff.shape)


if __name__ == '__main__':
	# Produce 2 different dataframes with the differences being on the row number
	df_main = produce_dataframe(dimensions=10, row_num=2)
	df_modified = produce_dataframe(dimensions=10, row_num=4)

	# Compare DataFrames to get differences and an updated df_modified with cells highlighted
	df_diff, df_modified_styled = compare_dataframes(df1=df_main, df2=df_modified)

	# Write DataFrames to excel workbook
	# Create an instance of excel
	with pd.ExcelWriter(path=FILE_PTH_OUTPUT, engine=excel_engine) as wkbk:
		# Write main data
		write_dataframe(workbook=wkbk, df=df_main, sheet_name='Raw Data')
		# Write modified data
		write_dataframe(workbook=wkbk, df=df_modified_styled, sheet_name='Modified Data', tab_color='green')
		# Write difference data
		write_dataframe(workbook=wkbk, df=df_diff, sheet_name='Difference Data', tab_color='blue')
