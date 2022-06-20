import pandas as pd
import numpy as np
import camelot.io as camelot

class ExtractPDFTables:

	# Initializing contructor variables
	def __init__(self, file_path, page_range=None, area=None, flavor=None, strip_text=None):
		self.file_path = file_path #string with file path
		self.page_range = page_range # defined page range with tables
		self.area = area # Table area, libraris like tabula allows to specify table area
		
		#Camelot parameters
		self.flavor = flavor
		self.strip_text = strip_text
	

	#Table Extrator with Camelot Library
	def getTablesCamelot(self):
		
		tables = camelot.read_pdf(self.file_path, 
								  pages=self.page_range, 
								  flavor=f'{self.flavor}', 
								  strip_text=f'{self.strip_text}')

		frames = []
		for i in tables:
			try:
				table = i.df
				frames.append(table)
			except:
				pass

		df =  pd.concat(frames)
		return df