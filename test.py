#!/usr/bin/env python
import unittest, merge_excel
import os
import json

class ProductTestCase(unittest.TestCase):

    def testCmdLineInput(self):
        os.system(
            'python merge_excel.py -d test_data -s Test01 -c "Server Type" -c "Item A" -c "Item C"'
            )
        targetData = [
                {
                    u'Server Type': u'Type 01',
                    u'Item A': 11,
                    u'Item C': 31,
                    u'Date': u'2015-11-04'
                },
                {
                    u'Server Type': u'Type 02',
                    u'Item A': 12,
                    u'Item C': 32,
                    u'Date': u'2015-11-04'
                },
                {
                    u'Server Type': u'Type 01',
                    u'Item A': 11,
                    u'Item C': 31,
                    u'Date': u'2015-11-05'
                },
                {
                    u'Server Type': u'Type 02',
                    u'Item A': 12,
                    u'Item C': 32,
                    u'Date': u'2015-11-05'
                }
                ]
        with open('output.json') as f:
            exData = json.loads(f.read())
            self.assertEqual(exData, targetData, 'Data not match' + str(exData))

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
        self.assertEqual(exData, targetData, 'Data not match')

    def testGetFileTime(self):
        excel_filename = '2015_11_04-test_excel_sheet.xlsx'
        pattern = r'[0-9]+'
        data_str = merge_excel.convert_file_date(
            filename = excel_filename, 
            pattern = pattern)
        self.assertEqual(data_str, '2015-11-04', 'Data error')

    def testMergeExcelData(self):
        excel_filename1 = 'test_data/2015_11_04-test_excel_sheet.xlsx'
        excel_filename2 = 'test_data/2015_11_04-test_excel_sheet.xlsx'
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
        targetData.extend(targetData)
        exData = merge_excel.bulk_merge_excel(
            filelist = [excel_filename1, excel_filename2],
            sheet='Test01',
            column=['Server Type', 'Item A', 'Item C']
            )
        self.assertEqual(exData, targetData, 'Data not match')

if __name__ == '__main__': unittest.main()
