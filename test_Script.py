import unittest
from Script import ExtractPDFTables
from pandas.testing import assert_frame_equal
import pandas as pd



class TestExtractPDFTables(unittest.TestCase):

	@classmethod
	def setUpClass(self):
		print('setup')
		self.table_1 = ExtractPDFTables('ESG-Frameworks/Mapping-Standards/SDG-GRI/sdg-gri.pdf', [list(range(3, 73)), list(range(74, 99))], area = [80.51, 90.42, 561.96, 814.18])
		self.table_2 = ExtractPDFTables('/desktop', list(range(20, 30)), area = [80.51, 90.42, 561.96, 814.18])

	@classmethod
	def tearDownClass(self):
		print('teardown')
		pass

    
	def test_filePath(self):
		print('test_1')
		self.assertEqual(self.table_1.file_path, 'ESG-Frameworks/Mapping-Standards/SDG-GRI/sdg-gri.pdf')

	def test_pageRange(self):
		print('Test_2')
		self.assertEqual(type(self.table_1.page_range), list)

	def test_getTablesSDG_GRI(self):
		print('Testing getTablesSDG_GRI...')

		expected = self.table_1.getTablesSDG_GRI()

		actual = pd.DataFrame(columns=['User_ID', 'UserName', 'Action'])

		assert_frame_equal(expected, actual, obj='DataFrame')


if __name__ == '__main__':
	unittest.main()