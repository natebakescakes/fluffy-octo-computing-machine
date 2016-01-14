import unittest, xlwt, xlrd, os
from ..scripts.build_out import build_out

class TestMakeWorksheetToTest(unittest.TestCase):

    def setUp(self):
        # Create Workbook
        master_files = {}
        wb = xlwt.Workbook()
        ws = wb.add_sheet('Test Sheet')
        wb.save('test.xls')

        # Instantiate worksheet objects
        master_files['xl_workbook'] = xlrd.open_workbook(os.path.join(os.getcwd(), 'test.xls'))
        master_files['xl_sheet_main'] = master_files['xl_workbook'].sheet_by_index(0)

    def tearDown(self):
        # Delete workbook
        os.remove('test.xls')

class TestMaximumLength(TestMakeWorksheetToTest):
    def test_maximum_length(self):
        assertEquals(0, 0, 'great stuff')

if __name__ == '__main__':
    unittest.main()
