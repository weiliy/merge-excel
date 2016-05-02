#!/usr/bin/env python
import unittest, merge_excel
from subprocess import call

class ProductTestCase(unittest.TestCase):

#   def testCmdLineInput(self):
#       call(['merge-excel', '-s Test01', '-c "Server Type"', '-c "Item A"', '-c "Item C"'])

    def testMergeExcel(self):
        excel_filename = 'test_data/2015_11_04-test_excel_sheet.xlsx'
        exData = merge_excel.extractExcel(filename=excel_filename, date='2015-11-04', sheet='Test01', column=['Server Type', 'Item A', 'Item C'])
        targetData = [
                {
                    'Server Type': u'Type 01',
                    'Item A': 11L,
                    'Item C': 31L,
                    'Date': '2015-11-04'
                },
                {
                    'Server Type': u'Type 02',
                    'Item A': 12L,
                    'Item C': 32L,
                    'Date': '2015-11-04'
                }
                ]
        # print 'target', targetData
        self.assertEqual(exData, targetData, 'Data not match')

    def testGetFileTime(self):
        excel_filename = '2015_11_04-test_excel_sheet.xlsx'
        pattern = r'[0-9]+'
        data_str = merge_excel.convert_file_date(
            filename = excel_filename, 
            pattern = pattern)
        self.assertEqual(data_str, '2015-11-04', 'Data error')

if __name__ == '__main__': unittest.main()
