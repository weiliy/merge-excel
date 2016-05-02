#!/usr/bin/env python
__metaclass__ = type
import sys, getopt
import openpyxl
import re
import os
import json

def extractExcel(filename, date, sheet, column):
    wb = openpyxl.load_workbook(filename = filename, read_only=True)
    ws = wb[sheet]

    headRowIdx = 0
    columnMap = { }
    resultData = [ ]

    for row in ws.rows:
        row_data = {}
        for cell in row:
            if cell.value == None:
                continue
            if headRowIdx == 0 or headRowIdx == cell.row:
                for col in column:
                    if cell.value == col:
                        headRowIdx = cell.row
                        # print 'Head row =', headRowIdx
                        columnMap[cell.column] = col 
                        # print 'Set', col, 'to column', cell.column
            else:
                if cell.column in columnMap.keys():
                    row_data[columnMap[cell.column]] = cell.value

        if row_data != {}:
            row_data['Date'] = date
            resultData.append(row_data)

    # print resultData
    return resultData

def convert_file_date(filename, pattern):
    return '-'.join(re.findall(pattern, filename))

def bulk_merge_excel(filelist, sheet, column):
    result_data = [ ]
    for filename in filelist:
        result_data.extend(
            extractExcel(
                filename = filename,
                sheet = sheet,
                column = column,
                date = convert_file_date(
                    filename = filename,
                    pattern = r'[0-9]+'
                    )
                )
            )
    return result_data

def main(argv):
    script_help_str = 'merge-excel.py -d <excel_direcoty> -s <sheet_name> -c <table_colume_name> [-c <table_colume_name>]...'
    input_dir = ''
    sheet_name = ''
    colume_list = [ ] 
    filelist = [ ]
    try:
        opts, args = getopt.getopt(argv,"hd:s:c:t:",["sheet=","colume="])
    except getopt.GetoptError:
        print script_help_str
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print script_help_str
            sys.exit()
        elif opt == "-d":
            input_dir = arg
        elif opt in ("-s", "--sheet"):
            sheet_name = arg
        elif opt in ("-c", "--colume"):
            colume_list.append(arg)

    print 'Input dir is ', input_dir
    print 'Sheet Name is ', sheet_name
    print 'Colume have ', colume_list

    print 'check files in', input_dir
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.xlsx'):
                filename = os.path.join(root, file)
                print '-', filename
                filelist.append(filename)

    data = bulk_merge_excel(
        filelist = filelist,
        sheet = sheet_name,
        column = colume_list
        )

    with open('output.json', 'wb') as f:
        f.write(
            json.dumps(
                data, 
                sort_keys=True,indent=4, separators=(',', ': ')
            )
        )

    

if __name__ == "__main__":
    main(sys.argv[1:])
