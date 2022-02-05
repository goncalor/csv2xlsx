#!/usr/bin/env python3

import csv
import sys
import os.path
import argparse
import xlsxwriter

CSV_SAMPLE_SIZE = 4096
MIN_COL_WIDTH = 8
MAX_COL_WIDTH = 90

parser = argparse.ArgumentParser(description="Convert CSV to XLSX")
parser.add_argument("csv_file",
                    metavar="<file.csv>",
                    type=str,
                    help="input .csv file")
parser.add_argument("--no-autofilter",
                    action="store_true",
                    help="do not apply autofilters to header row")
parser.add_argument("--no-freeze-header",
                    action="store_true",
                    help="do not freeze header row")
args = parser.parse_args()

if not os.path.isfile(args.csv_file):
    print("'{}' does not exist".format(args.csv_file))
    sys.exit(1)

csvfile = open(args.csv_file, newline='')
sample = csvfile.read(CSV_SAMPLE_SIZE)
csvfile.seek(0)
dialect = csv.Sniffer().sniff(sample)
has_header = csv.Sniffer().has_header(sample)

# Determine number of columns
reader = csv.reader(csvfile, dialect)
nr_columns = len(next(reader))
# Reset reader
csvfile.seek(0)

workbook = xlsxwriter.Workbook(os.path.splitext(args.csv_file)[0] + ".xlsx")
worksheet = workbook.add_worksheet()
format_row = workbook.add_format()
format_row.set_align("top")
col_width = [MIN_COL_WIDTH for c in range(nr_columns)]

for (r, row) in enumerate(reader):
    worksheet.write_row(r, 0, row, format_row)
    # Calculate column width
    for (c, coltxt) in enumerate(row):
        lentxt = len(coltxt)
        if lentxt > col_width[c]:
            col_width[c] = lentxt if lentxt < MAX_COL_WIDTH else MAX_COL_WIDTH

# Set column width
for (c, width) in enumerate(col_width):
    worksheet.set_column(c, c, width)

if not args.no_autofilter and has_header:
    worksheet.autofilter(0, 0, 0, nr_columns - 1)

if not args.no_freeze_header and has_header:
    worksheet.freeze_panes(1, 0)

workbook.close()
csvfile.close()
