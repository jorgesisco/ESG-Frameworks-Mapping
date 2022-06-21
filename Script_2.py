import pandas as pd
import numpy as np
import camelot.io as camelot

class ExtractPDFTables:

	# Initializing contructor variables
	def __init__(self, file_path, page_range=None, area=None, flavor=None, strip_text=None, get_codes=None):
		self.file_path = file_path #string with file path
		self.page_range = page_range # defined page range with tables
		self.area = area # Table area, libraris like tabula allows to specify table area
		
		#Camelot parameters
		self.flavor = flavor
		self.strip_text = strip_text

		self.get_codes = get_codes
    
    


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

		if self.get_codes != None:
			
			
			df = self.getCode(df, 
                            column=3,
                            regex='\d\d\d:',
                            delimiter='\n',
                            code_column=f'{self.get_codes[0]}_code',
                            remove_char=':')
			
			
			df = self.getCode(df, 
                            column=2,
                            regex='\w+.\d+\)',
                            delimiter='\n',
                            code_column=f'{self.get_codes[1]}_code',
                            remove_char=')')

			df = df.replace('', np.nan)
			df = df.dropna(how='all')
			df = df.fillna('')
			return df
    
		else:
			df = df.replace('', np.nan)
			df = df.dropna(how='all')
			df = df.fillna('')
 
			return df



    
	def getCode(self, df,
                    column,
                    regex, 
                    delimiter, 
                    code_column,
                    remove_char=None):
                    
                    df[code_column] = df[column].str.findall(regex).apply(set).str.join(delimiter)

                    if remove_char != None:
                        df[code_column] = df[code_column].str.replace(remove_char, '', regex=False)

                    return df