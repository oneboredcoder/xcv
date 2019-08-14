import os
import sys
import numpy.random.common # pyinstaller looses the path to this library
import numpy.random.bounded_integers # pyinstaller looses the path to this library
import numpy.random.entropy # pyinstaller looses the path to this library
import pandas

# install required modules
# pip install xlrd xlwt pandas

from datetime import date

today = date.today()
today_str = today.strftime("%d%m%Y")

def main():
	# read .xls and save as .csv with required columns only
	wb = pandas.read_excel(sys.argv[1], usecols=['column_1','column_2'])
	wb.to_csv("temp.csv", index=False, sep=";", encoding="utf-8")

	# read temp.csv
	df = pandas.read_csv("temp.csv", sep=";")

	# fix lowercase letters
	df['column_1'] = df['column_1'].str.upper()
	df['column_2'] = df['column_2'].str.upper()

	# strip whitespaces
	df['column_1'] = df['column_1'].str.strip()
	df['column_2'] = df['column_2'].str.strip()
	df.to_csv("temp.csv", index=False, sep=";", encoding="utf-8")

	# find incorrect entries in column_2
	problems = df[~(df['column_2'].str.len() == 3)]

	# if there are no problems, generate final .csv
	if problems.empty:
		df.to_csv("column_1_column_2_" + today_str + ".csv", index=False, sep=";", encoding="utf-8")
		# and delete temporary file
		os.remove("temp.csv")
	else:
		print("Problems detected")
		print(problems)

if __name__ == "__main__":
	main()
