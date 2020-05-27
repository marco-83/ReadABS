#  Functions to locate tables in ABS spreadsheets and save location and descriptions to class TableData

import pandas as pd
import copy
import xlsxwriter
import re
from operator import itemgetter

class TableData:
    def __init__(self, sheet_name):
        self.sheet_name = sheet_name
        self.last_row_in_sheet = None
        self.rows = set()
        self.cols = set()
        self.indentation_levels = {}
        self.columns_with_indentation = []
        self.top_row = None
        self.top_header_row = None
        self.table_type = None
        self.row_descriptions = []
        # self.non_blank_rows_above_data = set()
        self.column_header_locations = set()
        self.column_titles = {}
        self.row_titles = {}
        self.row_descriptions_header_row = None
        self.merged_meta_data = []
        self.extra_meta_data = set()
        self.merged_meta_data_row_headings = []
        self.table_completed = False

    def add_row(self, row):
        self.rows.add(row)

    def add_col(self, col):
        self.cols.add(col)

    def __repr__(self):
        return "TableData sheet_name:% s last_row_in_sheet:% s table_type:% s rows:% s cols:% s " \
               "indentation_levels:% s columns_with_indentation:% stop_row:% s top_header_row:% s " \
               "row_descriptions:% s column_header_locations:% s column_titles:% s row_titles:% s " \
               "merged_meta_data:% s extra_meta_data:% s merged_meta_data_row_headings:% s" % \
               (self.sheet_name, self.last_row_in_sheet, self.table_type, self.rows, self.cols,
                self.indentation_levels, self.columns_with_indentation, self.top_row, self.top_header_row,
                self.row_descriptions, self.column_header_locations, self.column_titles, self.row_titles,
                self.merged_meta_data, self.extra_meta_data, self.merged_meta_data_row_headings)

    def to_dict(self):
        return {
            'Tab': self.sheet_name,
            'Rows': str(min(self.rows)) + ':' + str(max(self.rows)),
            'Columns': xlsxwriter.utility.xl_col_to_name(min(self.cols)) + ':' +
                       xlsxwriter.utility.xl_col_to_name(max(self.cols)),
        }

BUILTIN_FORMATS = {
    0: 'General',
    1: '0',
    2: '0.00;0.0',
    #2: '0.00);(0.0)',
    3: '#,##0',
    4: '#,##0.00;#,##0.0',
    #4: '#,##0.00);(#,##0.0)',
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',

    37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)',
    40: '#,##0.00_);[Red](#,##0.00)',

    41: r'_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
    42: r'_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
    43: r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',

    44: r'_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
    50: '0.0',
    51: '#,##0.0'}

BUILTIN_FORMATS_MAX_SIZE = 164
BUILTIN_FORMATS_REVERSE = dict(
    [(value, key) for key, value in BUILTIN_FORMATS.items()])

FORMAT_GENERAL = BUILTIN_FORMATS[0]
FORMAT_TEXT = BUILTIN_FORMATS[49]
FORMAT_NUMBER = BUILTIN_FORMATS[1]
FORMAT_NUMBER_0 = BUILTIN_FORMATS[50]
FORMAT_NUMBER_00 = BUILTIN_FORMATS[2]
FORMAT_NUMBER_COMMA_SEPARATED1 = BUILTIN_FORMATS[4]
FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-'
FORMAT_NUMBER_COMMA_SEPARATED3 = BUILTIN_FORMATS[51]
FORMAT_PERCENTAGE = BUILTIN_FORMATS[9]
FORMAT_PERCENTAGE_00 = BUILTIN_FORMATS[10]
FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd'
FORMAT_DATE_YYMMDD = 'yy-mm-dd'
FORMAT_DATE_DDMMYY = 'dd/mm/yy'
FORMAT_DATE_DMYSLASH = 'd/m/y'
FORMAT_DATE_DMYMINUS = 'd-m-y'
FORMAT_DATE_DMMINUS = 'd-m'
FORMAT_DATE_MYMINUS = 'm-y'
FORMAT_DATE_XLSX14 = BUILTIN_FORMATS[14]
FORMAT_DATE_XLSX15 = BUILTIN_FORMATS[15]
FORMAT_DATE_XLSX16 = BUILTIN_FORMATS[16]
FORMAT_DATE_XLSX17 = BUILTIN_FORMATS[17]
FORMAT_DATE_XLSX22 = BUILTIN_FORMATS[22]
FORMAT_DATE_DATETIME = 'yyyy-mm-dd h:mm:ss'
FORMAT_DATE_TIME1 = BUILTIN_FORMATS[18]
FORMAT_DATE_TIME2 = BUILTIN_FORMATS[19]
FORMAT_DATE_TIME3 = BUILTIN_FORMATS[20]
FORMAT_DATE_TIME4 = BUILTIN_FORMATS[21]
FORMAT_DATE_TIME5 = BUILTIN_FORMATS[45]
FORMAT_DATE_TIME6 = BUILTIN_FORMATS[21]
FORMAT_DATE_TIME7 = 'i:s.S'
FORMAT_DATE_TIME8 = 'h:mm:ss@'
FORMAT_DATE_TIMEDELTA = '[hh]:mm:ss'
FORMAT_DATE_YYMMDDSLASH = 'yy/mm/dd@'
FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
FORMAT_CURRENCY_USD = '$#,##0_-'
FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'
COLORS = r"\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW)\]"
LITERAL_GROUP = r'"[^"]+"'
LOCALE_GROUP = r'\[\$[^\]]+\]'
STRIP_RE = re.compile("{0}|{1}|{2}".format(COLORS, LITERAL_GROUP, LOCALE_GROUP), re.IGNORECASE + re.UNICODE)


def is_date_format(fmt):
    if fmt is None:
        return False
    fmt = fmt.split(";")[0]  # only look at the first format
    fmt = STRIP_RE.sub("", fmt)
    return re.search("[dmhysDMHYS]", fmt) is not None


# def is_number_format(fmt):  # includes currency and percentages
#     if fmt is None:
#         return False
#     fmt = fmt.split(";")[0]  # only look at the first format
#     fmt = BUILTIN_FORMATS_REVERSE.get(fmt, '')
#     if fmt in [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 37, 38, 39, 40, 41, 42, 43, 44, 48, 50, 51]:
#         return True
#     else:
#         return False

def is_numeric_format(fmt):
    if fmt is None:
        return False
    fmt = fmt.split(";")[0]  # only look at the first format
    fmt = STRIP_RE.sub("", fmt)
    return re.search("0", fmt) is not None


def is_numeric(value):
    if value is None:
        return False
    if isinstance(value, float):
        return True
    elif isinstance(value, int):
        return True
    else:
        return False


def is_builtin(fmt):
    return fmt in BUILTIN_FORMATS.values()


def import_spreadsheet(excel_workbook="81670do001_201718.xls", filter_tabs=True):
    # TODO: Block macros
    xl_workbook = excel_workbook
    data_sheets = []
    if filter_tabs:
        for sheet in xl_workbook.sheetnames:
            if sheet.startswith("Table") or sheet.startswith("Data"):
                data_sheets.append(sheet)
    else:
        for sheet in xl_workbook.sheetnames:
            data_sheets.append(sheet)

    return xl_workbook, data_sheets


def define_table(xl_workbook, data_sheets, allowed_blank_rows, spreadsheet_type):
    # Initiate tables with data
    tables = locate_data(xl_workbook, data_sheets, allowed_blank_rows, spreadsheet_type)

    # Add row descriptions and column headers
    for table in tables:
        try:

            # Extract some details from the table
            first_col = min(table.cols)
            last_col = max(table.cols)
            first_row = min(table.rows)
            last_row = max(table.rows)
            data_cols = table.cols

            table.last_row_in_sheet = find_last_row_in_sheet(xl_workbook, table.sheet_name)

            # Find row descriptions
            table.row_descriptions, table.table_type = locate_row_descriptions(xl_workbook, first_col, first_row,
                                                                                    table.sheet_name, table.cols)


            # How many levels of indentation are there, if any, in the descriptor column?
            # If there are more than 1 row descriptor columns, then we can assume there is no indentation
            #if len(table.row_descriptions) == 1:
            table.indentation_levels, table.columns_with_indentation, table.top_row, \
            table.top_header_row = describe_indentation(table.rows, table.cols, table.sheet_name,
                                                        table.row_descriptions, xl_workbook)
            #else:
            table.merged_meta_data_row_headings, table.row_descriptions_header_row, \
            table.row_titles = describe_row_headers(table.rows, table.cols, table.sheet_name, table.row_descriptions,
                                                    table.columns_with_indentation, xl_workbook)
            # Find column headings
            if spreadsheet_type == "Time series":
                table.column_titles, \
                table.column_header_locations = describe_col_headings_timeseries(xl_workbook, table.sheet_name,
                                                                                      table.rows, table.cols,
                                                                                      table.top_header_row)
            else:
                table.merged_meta_data, table.column_header_locations, \
                table.extra_meta_data = describe_col_headings(xl_workbook, table.sheet_name,
                                                              table.rows, table.cols, table.row_descriptions,
                                                              table.top_row, table.top_header_row,
                                                              table.row_descriptions_header_row, spreadsheet_type,
                                                              table.last_row_in_sheet)

        except (ValueError, TypeError):
            table.table_completed = False
        else:
            table.table_completed = True

    return tables


def find_last_row_in_sheet(xl_workbook, sheet_name):
    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    for r, row in enumerate(sheet.iter_rows(min_row=1)):
        for c, cell_obj in enumerate(row):
            if cell_obj.value:
                last_row_in_sheet = r

    return last_row_in_sheet


def locate_data(xl_workbook, data_sheets, allowed_blank_rows, data_type):
    """ Function to locate the data in the spreadsheet and assign it to a TableData class """

    # Initiate the first table
    tables = []
    table_number = -1

    for s in data_sheets:
        sheet = xl_workbook.get_sheet_by_name(s)
        found_table = False
        r = 0
        found_data = False
        looking_for_multiple_tables = False
        blank_row = False
        blank_row_count = 0
        quit_loop = False
        date_cols = []

        if data_type == "Time series":
            for i, row in enumerate(sheet.iter_rows()):
                # for i in range(sheet.nrows):
                #  row = sheet.row(i)
                if quit_loop:
                    break
                for idx, cell_obj in enumerate(row):
                    if cell_obj.value == "Series ID":
                        start_row = i + 1
                        quit_loop = True
                        break
        else:
            start_row = 1

        quit_loop = False

        if 'start_row' not in locals():
            start_row = 1

        # Find data
        for i, row in enumerate(sheet.iter_rows(min_row=start_row)):  # Is the first row 0 or 1?
            # if i < 9:
            #     continue
            if quit_loop:
                break
            # row = sheet.row(i)
            c = 0
            if blank_row:
                blank_row_count += 1
            else:
                blank_row_count = 0
            if blank_row_count >= allowed_blank_rows:
                found_data = False
                found_table = False
                blank_row = False
                date_cols = []
                continue
            if found_data:
                looking_for_multiple_tables = True
                blank_row = True  # Temporary assignment, will be made false if something is found in the row
                blank_row_count = 0
                r += 1
            found_data = False
            for idx, cell_obj in enumerate(row):
                if isinstance(cell_obj.value, str):
                    # if "STANDARD ERROR" in cell_obj.value:
                    #     quit_loop = True
                    #     break
                    if not found_data:
                        if cell_obj.value in ["Year", "year", "Years", "Date", "Month", "Day"]:
                            date_cols.append(idx + 1)
                            continue
                cell_type = cell_obj.number_format
                # Check if it is a total row (p.s. this is very specific, sometimes total row is at the top and bold
                if is_numeric(cell_obj.value) and not found_data and idx != 0:  # Top left data cell
                    # left_of_data = sheet.row(i)[idx - 1].value
                    left_of_data = row[idx-1].value
                if is_numeric(cell_obj.value) and (not cell_obj.font.bold or left_of_data == "Total") and idx + 1 not in date_cols:

                    found_data = True
                    if not found_table:
                        tables.append(TableData(sheet_name=sheet.title))
                        table_number += 1
                        found_table = True
                    tables[table_number].add_row(i + start_row)
                    tables[table_number].add_col(idx + 1)
                    c += 1
                # Sometimes there is a 'total' row that is bold. Let's not skip it.
                # If the table has been started, allow bold rows.
                # p.s this might be redundant if the total row is always called 'Total' as per above if statement
                if found_table and is_numeric(cell_obj.value) and cell_obj.font.bold:
                    found_data = True
                    tables[table_number].add_row(i + start_row)
                    tables[table_number].add_col(idx + 1)
                    c += 1
                if cell_obj.value and looking_for_multiple_tables is True:
                    blank_row = False
                    blank_row_count = 0

    return tables


def locate_row_descriptions(xl_workbook, first_col, first_row, sheet_name, data_cols):
    """ Function to locate the column where the row headers are. Returns a single column """

    if len(data_cols) == 1:
        table_type = "long format"
    elif len(data_cols) > 1:
        table_type = "wide format"
    else:
        table_type = "no data"
    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    row_descriptions = []

    # Where is the descriptor column?
    if table_type == "wide format":
        for c in range(1, first_col):
            cell = sheet.cell(row=first_row, column=c)
            if is_numeric(cell.value) and not is_date_format(cell.number_format):
                continue
            else:
                row_descriptions.append(c)
                # break
    elif table_type == "long format":
        for c in range(1, first_col):
            cell = sheet.cell(row=first_row - 1, column=c)
            if is_numeric(cell.value) and not is_date_format(cell.number_format):

                continue
            else:
                row_descriptions.append(c)

    if len(row_descriptions) == 0:
        print("Something went wrong. Could not locate row headings")
        return "Could not locate row headings", table_type  # Possibly a pivot table

    return row_descriptions, table_type


def describe_indentation(data_rows, data_cols, sheet_name, row_descriptions, xl_workbook):
    """ Function to find row descriptions. If they are indented (subcategories) then store that information so multiple
     row desciptors can be made """

    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    first_row = min(data_rows)
    last_row = max(data_rows)
    top_cell_row = min(data_rows)
    first_col = min(data_cols)

    # Identify the row(s) above the data that describe the columns.
    # Starts 1 row above the data, working up. If a blank row is found, then stop (unless no column headers have been \
    # identified yet).
    rows_above_data = range(1, first_row)
    non_blank_rows_above_data = set()
    blank_row = False
    found_a_non_blank_row = False
    for r in reversed(rows_above_data):
        if blank_row and found_a_non_blank_row:
            break
        row = sheet[r]
        for idx, cell_obj in enumerate(row):
            if idx >= first_col - 1:
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if not cell_obj.value:
                    blank_row = True
                else:
                    blank_row = False
                    found_a_non_blank_row = True
                    non_blank_rows_above_data.add(r)
                    break
    top_header_row = min(non_blank_rows_above_data)

    # How many levels of indentation are there, if any?
    #indentation_levels = set()
    indentation_levels = {}
    start_row = max(non_blank_rows_above_data)

    # Start 2 row above first row of data
    for i, c in enumerate(row_descriptions):
        indentation_levels[c] = set()
        for r in range(start_row, last_row + 1):
            indentation = sheet.cell(row=r, column=row_descriptions[i]).alignment.indent
            if sheet.cell(row=r, column=row_descriptions[i]).value:
                indentation_levels[c].add(int(indentation))

    columns_with_indentation = []
    for c in indentation_levels:
        if len(indentation_levels[c]) > 1:
            indentation_levels[c] = list(indentation_levels[c])
            indentation_levels[c].sort()
            columns_with_indentation.append(c)

    for r in data_rows:
        for idx, c in enumerate(columns_with_indentation):
            for i in indentation_levels[c]:
                cell_row = r
                indentation_cell = sheet.cell(cell_row, columns_with_indentation[idx]).alignment.indent
                while indentation_cell > i:
                    cell_row -= 1
                    indentation_cell = sheet.cell(cell_row, columns_with_indentation[idx]).alignment.indent
                    if cell_row < top_cell_row:
                        top_cell_row = cell_row

    return indentation_levels, columns_with_indentation, top_cell_row, top_header_row


def describe_row_headers(data_rows, data_cols, sheet_name, row_descriptions, columns_with_indentation, xl_workbook):
    """ Function to find row descriptions. Only run if there is more than 1 row descriptor column,
    and check for merged cells"""

    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    first_row = min(data_rows)
    top_cell_row = min(data_rows)
    first_col = min(data_cols)
    other_columns = set(i for i in row_descriptions if i not in columns_with_indentation)

    # Identify the row(s) above the data that describe the columns.
    # Starts 1 row above the data, working up. If a blank row is found, then stop (unless no column headers have been \
    # identified yet).
    rows_above_data = range(1, first_row)

    non_blank_rows_above_data = set()
    blank_row = False
    found_a_non_blank_row = False
    for r in reversed(rows_above_data):
        if blank_row and found_a_non_blank_row:
            break
        row = sheet[r]
        for idx, cell_obj in enumerate(row):
            if idx >= first_col:
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if cell_obj.value == "":
                    blank_row = True
                else:
                    blank_row = False
                    found_a_non_blank_row = True
                    non_blank_rows_above_data.add(r + 1)
                    break
    #top_header_row = min(non_blank_rows_above_data)

    # In xlrd:
    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c [0,2] is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d [1,3] is the bottom right coordinate (row / col, starting with 1) where the merge finishes (who knows why?)

    # In openpyxl:
    # sheet.merged_cells.ranges returns a list of 'CellRange' objects.
    # Use the '.bounds' attribute to extract the coordinates.
    # Each tuple has 4 elements a,b,c,d
    # b,a [1,0] is the top-left coordinate (row / col, starting with 1) where the merge starts.
    # d,c [3,2] is the bottom right coordinate (row / col, starting with 1) where the merge finishes
    # Or:
    # b,d [1,3] are the row coordinates (from / to, starting with 1)
    # a,c [0,2] are the column coordinates (from / to, starting with

    assert isinstance(sheet.merged_cells, object)
    all_mergers_ranges = sheet.merged_cells.ranges
    all_mergers = []
    for i in all_mergers_ranges:
        all_mergers.append(i.bounds)

    merged_meta_data_row_headings = []
    merged_meta_data_col_headings = []
    for i in all_mergers:
        # Only keep the merged cells that are above the data; not to the left and not to the right
        if i[1] in data_rows and i[0] < first_col:
            merged_meta_data_row_headings.append(i)
        elif i[0] < first_col:
            merged_meta_data_col_headings.append(i)

    merged_meta_data_col_headings = [a_tuple[1] for a_tuple in merged_meta_data_col_headings]

    # Look for a header row above the row descriptions.
    # The header row must have something in each cell above the row descriptions, otherwise return None.
    row_descriptions_header_row = None
    for r in reversed(range(1, min(data_rows))):  # Note: 'range' does not include end point
        if r not in merged_meta_data_col_headings and r > top_cell_row - 2:  # 2 above data is a bit arbitrary
            if row_descriptions_header_row:
                break
            for c in other_columns:
                if not sheet.cell(row=r, column=c).value:
                    row_descriptions_header_row = None
                    continue
                else:
                    row_descriptions_header_row = r

    row_titles = {}
    if row_descriptions_header_row:
        for i, c in enumerate(other_columns):
            if sheet.cell(row=row_descriptions_header_row, column=c).value:
                row_titles[c] = sheet.cell(row=row_descriptions_header_row, column=c).value
            else:
                row_titles[c] = "Row_description_title_sub_" + str(i)

    else:
        for i, c in enumerate(other_columns):
            row_titles[c] = "Row_description_title_sub_" + str(i)

    return merged_meta_data_row_headings, row_descriptions_header_row, row_titles


def describe_col_headings_timeseries(xl_workbook, sheet_name, data_rows, data_cols, top_header_row):
    first_row = min(data_rows)
    first_col = min(data_cols)
    sheet = xl_workbook.get_sheet_by_name(sheet_name)

    rows_above_data = range(1, first_row)
    quit_loop = False
    for i, row in enumerate(sheet.iter_rows()):
        if quit_loop:
            break
        for idx, cell_obj in enumerate(row):
            if cell_obj.value == "Series ID":
                series_id_position = [row, idx + 1]  # location of "Series ID" field
                quit_loop = True
                break

    if 'series_id_position' not in locals():
        series_id_position = [9, 0]

    column_header_locations = set()
    for r in range(1, first_row):
    #for r in range(top_header_row, first_row):
        if sheet.cell(row=r, column=first_col).value is not None:
            column_header_locations.add(r)
    column_titles = {}
    for r in column_header_locations:
        column_titles[r] = sheet.cell(row=r, column=series_id_position[1]).value
        if not column_titles[r]:
            column_titles[r] = "Description"
    return column_titles, column_header_locations


def describe_col_headings(xl_workbook, sheet_name, data_rows, data_cols, row_descriptions, top_row, top_header_row,
                          row_descriptions_header_row, spreadsheet_type, last_row_in_sheet):
    """ Find column headings. There might be multiple column headings (above each other) that might be units ($, %, etc)
    or they might be merged cells """

    first_row = min(data_rows)
    last_col = max(data_cols)
    first_col = min(data_cols)
    sheet = xl_workbook.get_sheet_by_name(sheet_name)

    # In xlrd:
    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c [0,2] is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d [1,3] is the bottom right coordinate (row / col, starting with 1) where the merge finishes (who knows why?)

    # In openpyxl:
    # sheet.merged_cells.ranges returns a list of 'CellRange' objects.
    # Use the '.bounds' attribute to extract the coordinates.
    # Each tuple has 4 elements a,b,c,d
    # b,a [1,0] is the top-left coordinate (row / col, starting with 1) where the merge starts.
    # d,c [3,2] is the bottom right coordinate (row / col, starting with 1) where the merge finishes
    # Or:
    # b,d [1,3] are the row coordinates (from / to, starting with 1)
    # a,c [0,2] are the column coordinates (from / to, starting with )

    assert isinstance(sheet.merged_cells, object)
    all_mergers_ranges = sheet.merged_cells.ranges
    all_mergers = []
    for i in all_mergers_ranges:
        all_mergers.append(i.bounds)

    rows_above_data = range(1, first_row)

    column_header_locations = set()
    blank_row = False
    found_a_non_blank_row = False
    found_a_column_header = False

    def check_rows(idx, merged_meta_data):
        if merged_meta_data:
            merged_col1 = list(zip(*merged_meta_data))[0]
            if idx >= first_col or idx in merged_col1:
                return True
            else:
                return False
        else:
            if idx >= first_col:
                return True
            else:
                return False

    blank_row = False
    found_a_non_blank_row = False
    for r in reversed(rows_above_data):
        if blank_row and found_a_non_blank_row and found_a_column_header:
            break
        row = sheet[r]
        merged_meta_data = list(filter(lambda x: x[1] == r, all_mergers))
        for idx, cell_obj in enumerate(row):
            cell_type = cell_obj.number_format
            # If the column index is greater than the first data column or is in the first column of merged cells
            if check_rows(idx + 1, merged_meta_data):
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if cell_obj.value == "" or not cell_obj.value:
                    blank_row = True
                # At least one of the column headers must be text or bold for this to work
                elif (cell_type == "General" or cell_obj.font.bold) and cell_obj.value:
                    blank_row = False
                    found_a_non_blank_row = True
                    found_a_column_header = True
                    column_header_locations.add(r)
                    break
                else:
                    blank_row = False
                    found_a_non_blank_row = True

    # column_header_locations.discard(top_row)

    # Extra meta data above top left cell
    extra_meta_data = set()
    if spreadsheet_type == "Data cube":
        if row_descriptions_header_row:
            column_headers_already_included = {top_row, row_descriptions_header_row}
        else:
            column_headers_already_included = {top_row}

        if all_mergers:
            for r in column_header_locations:
                mergers_filtered = [tup for tup in all_mergers if tup[1] == r]
                if mergers_filtered:
                    for c in row_descriptions:
                        if c in list(zip(*mergers_filtered))[0]:
                            column_headers_already_included.add(r)

        rows_above_data = range(1, top_header_row + 1)
        rows_above_data = list(filter(lambda i: i not in column_headers_already_included, [*rows_above_data]))

        blank_row = False
        for r in reversed(rows_above_data):
            if blank_row:
                break
            row = sheet[r]
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx + 1 in row_descriptions:
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "" or not cell_obj.value:
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif cell_obj.number_format == "General" and cell_obj.value:
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False

        blank_row = False
        for r in rows_above_data:
            if blank_row:
                break
            row = sheet[r]
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx + 1 in row_descriptions:
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "" or not cell_obj.value:
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif cell_obj.value:
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False

    else:
        columns_to_evaluate = range(1, max(row_descriptions))
        rows_above_top_header_row = range(1, top_header_row)
        column_headers_already_included = {top_header_row}
        if all_mergers:
            for r in column_header_locations:
                mergers_filtered = [tup for tup in all_mergers if tup[1] == r]
                if mergers_filtered:
                    for c in row_descriptions:
                        if c in list(zip(*mergers_filtered))[0]:
                            column_headers_already_included.add(r)
                # for c in columns_to_evaluate:
                #     if r in list(zip(*all_mergers))[1] and c in list(zip(*all_mergers))[0]:
                #         column_headers_already_included.add(r)

        blank_row = False
        rows_above_data = list(filter(lambda i: i not in column_headers_already_included, [*rows_above_top_header_row]))
        for r in reversed(rows_above_data):
            row = sheet[r]
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx in columns_to_evaluate:
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "" or not cell_obj.value:
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif cell_obj.number_format == "General" and cell_obj.value:
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False

    other_rows = set(i for i in range(last_row_in_sheet + 1) if i not in data_rows and i > max(column_header_locations))
    all_rows = column_header_locations.union(other_rows)

    merged_meta_data = []
    #rows_with_merged_meta_data = set()
    for i in all_mergers:
        # Only keep the merged cells that are above the data; not to the left and not to the right
        # if i[0] in column_headers_and_rows_above_data and \
        #        i[2] >= first_col and i[1] - 1 in column_headers_and_rows_above_data and i[3] - 1 <= last_col:
        if i[1] in all_rows and i[0] <= last_col:
            merged_meta_data.append(i)
            #rows_with_merged_meta_data.add(i[1])
    return merged_meta_data, column_header_locations, extra_meta_data


def merged_data_function(xl_workbook, sheet_name, merged_data_cols, data_cols, data_rows,
                              extra_rows, last_row_in_sheet, spreadsheet_type, column_header_locations,
                              column_position=1):
    """ Function to extract data from merged cells
    merged_data_cols is a list of tuples. Each tuple is in the format used by xlrd function merged_cells """

    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    last_row = max(data_rows)
    other_rows = set(i for i in range(last_row_in_sheet + 1) if i not in data_rows and i > max(column_header_locations))
    all_rows = column_header_locations.union(other_rows)

    column_headings = pd.DataFrame()
    merged_meta_data = list(filter(lambda x: x[1] in all_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    other_rows = set(i for i in range(last_row) if i not in data_rows and i > max(column_header_locations) - 1)
    all_rows = column_header_locations.union(other_rows)

    merged_meta_data_subheadings_potential = list(filter(lambda x: x[1] in other_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    merged_meta_data_cols = [el[0:3:2] for el in merged_meta_data_subheadings_potential]
    duplicates = list(set([ele for ele in merged_meta_data_cols
                           if merged_meta_data_cols.count(ele) > 1]))

    merged_meta_data_subheadings = []
    for i in merged_meta_data_subheadings_potential:
        for j in duplicates:
            if i[0] == j[0] and i[2] == j[1]:
                merged_meta_data_subheadings.append(i)

    # duplicate_rows = set(el[0] for el in merged_meta_data_subheadings)
    #
    # merged_meta_data_subheadings = []
    # for i in merged_meta_data:
    #     for j in duplicates:
    #         if i[0] == j[0] and i[2] == j[1]:
    #             merged_meta_data_subheadings.append(i)

    # Remove the subheading rows
    merged_meta_data = [x for x in merged_meta_data if x not in merged_meta_data_subheadings]
    subheading_rows = [i[1] for i in merged_meta_data_subheadings]
    rows_not_subheadings = [x for x in column_header_locations if x not in subheading_rows]

    # Find how the merged data relates to the columns
    first_col = min(data_cols)
    all_positions = []
    for i in rows_not_subheadings:
        for j in data_cols:
            #all_positions.append((i, i + 1, j, j + 1))
            #all_positions.append((j, i, j + 1, i + 1))
            all_positions.append((j, i, j, i))

    all_merged_positions = []
    for i in merged_meta_data:
        j = i[0]  # start column  2
        k = 0
        while j <= i[2]:   # 2 <= 4
            for cells in range(i[0], i[0] + k + 1):
                all_merged_positions.append((cells, i[1], i[0] + k, i[3])) # ((i[0], i[1], cells, i[2] + k))
                #all_merged_positions.append((i[1], i[3], cells, i[0] + k))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(1, 0))

    values = [0]
    values.extend(col for col in range(min(data_cols), max(data_cols)) if col not in data_cols)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_cols = dict(zip(keys, values))

    column_titles = {}
    if spreadsheet_type == "Census TableBuilder":
        i = 1
        for r in column_header_locations:
            column_titles[r] = sheet.cell_value(rowx=r, colx=0)
            if column_titles[r] == '':
                column_titles[r] = "Column_description_title_" + str(i)
                i += 1
    else:
        for i, r in enumerate(rows_not_subheadings):
            column_titles[r] = "Row_description_title_" + str(i)

    for r in rows_not_subheadings:
        column_heading = column_titles[r]
        for i in merged_meta_data_extended:
            if i[1] == r:
                row_position = i[0] - first_col + 1  # row position in df (in other words, the column number)
                # Filter out entries that occur in empty columns
                empty_cols_filtered = dict(filter(lambda elem: elem[1] <= i[3], empty_cols.items())) #TODO: <=i[3]?
                if empty_cols_filtered:
                    row_position = row_position - max(empty_cols_filtered, key=empty_cols_filtered.get)
                for k in range(i[0], i[2] + 1):
                    cell = sheet.cell(row=i[1], column=i[0])
                    if row_position >= 1:
                        if is_date_format(cell.number_format):
                            column_headings.loc[row_position, column_heading] = pd.to_datetime(cell.value).\
                                strftime('%d/%m/%Y')
                        else:
                            column_headings.loc[row_position, column_heading] = cell.value
                    row_position += 1

    columns_to_evaluate = [1]  # Assume the extra row info is all in column A
    for j in extra_rows:
        column_heading = 'Col_desc_' + str(column_position)
        for i in columns_to_evaluate:
            cell = sheet.cell(row=j, column=i)
            for row_position in range(1, len(data_cols) + 1):
                if is_date_format(cell.number_format):
                    column_headings.loc[row_position, column_heading] = pd.to_datetime(cell.value).strftime('%d/%m/%Y')
                else:
                    column_headings.loc[row_position, column_heading] = cell.value
        column_position += 1

    return column_headings


def merged_data_row_headings_function(xl_workbook, sheet_name, merged_data_rows, data_rows, row_descriptions,
                                      row_titles, top_header_row):
    """ Function to extract data from merged cells
    merged_data_rows is a list of tuples. Each tuple is in the format used by xlrd function merged_cells """

    sheet = xl_workbook.get_sheet_by_name(sheet_name)

    other_rows = set(i for i in range(top_header_row + 1, max(data_rows)) if i not in data_rows)
    all_rows = data_rows.union(other_rows)


    # In openpyxl:
    # sheet.merged_cells.ranges returns a list of 'CellRange' objects.
    # Use the '.bounds' attribute to extract the coordinates.
    # Each tuple has 4 elements a,b,c,d
    # b,a [1,0] is the top-left coordinate (row / col, starting with 1) where the merge starts.
    # d,c [3,2] is the bottom right coordinate (row / col, starting with 1) where the merge finishes
    # Or:
    # b,d [1,3] are the row coordinates (from / to, starting with 1)
    # a,c [0,2] are the column coordinates (from / to, starting with 1)

    row_headings = pd.DataFrame()
    all_positions = []
    for i in all_rows: #data_rows:
        for j in row_descriptions:
            #all_positions.append((i, i + 1, j, j + 1))
            all_positions.append((j, i, j, i))
            #all_positions.append((i, j, i, j))

    merged_meta_data_row_headings = list(filter(lambda x: x[1] in data_rows, merged_data_rows))

    all_merged_positions = []
    for i in merged_meta_data_row_headings:
        j = i[1]  # start row
        k = 0
        while j <= i[3]:
            for cells in range(i[1], i[1] + k + 1):
                all_merged_positions.append((i[0], cells, i[2], i[1]+k)) #((cells, i[0] + k, i[2], i[3]))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data_row_headings)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(1, 0))

    values = [0]
    values.extend(row for row in range(min(data_rows), max(data_rows)+1) if row not in data_rows)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_rows = dict(zip(keys, values))

    # Look for a header row above the row descriptions.
    # The header row must have something in each cell above the row descriptions, otherwise return None.
    row_descriptions_header_row = None
    for r in reversed(range(1, min(data_rows))):  # Note: 'range' does not include end point
        if row_descriptions_header_row:
            break
        for c in row_descriptions:
            if not sheet.cell(row=r, column=c).value:
                row_descriptions_header_row = None
                continue
            else:
                row_descriptions_header_row = r

    descriptions_in_other_rows = []
    columns_included = set()
    first_data_row = min(data_rows)
    for c in row_descriptions:
        column_heading = row_titles[c]
        for i in merged_meta_data_extended:
            if i[0] == c:
                row_position = i[1] - first_data_row  # row position in df (in other words, the row number)
                # Filter out entries that occur in empty columns
                empty_rows_filtered = dict(filter(lambda elem: elem[1] < i[3], empty_rows.items()))
                if empty_rows_filtered:
                    row_position = row_position - max(empty_rows_filtered, key=empty_rows_filtered.get)
                    for k in range(i[1], i[3] + 1):
                        cell_value = sheet.cell(row=i[1], column=i[0]).value
                        if cell_value:
                            descriptions_in_other_rows.append({'Row': i[1], 'Col': i[0],
                                                               'row_position': row_position,
                                                               'Desc_row'+str(i[0]): cell_value})
                for k in range(i[1], i[3] + 1):
                    cell = sheet.cell(row=i[1], column=i[0])

                    if row_position >= 0:
                        if is_date_format(cell.number_format):
                            row_headings.loc[row_position, column_heading] = pd.to_datetime(cell.value).strftime('%d/%m/%Y')
                        else:
                            row_headings.loc[row_position, column_heading] = cell.value
                        if cell.value and i[1] in data_rows:
                            columns_included.add(i[0])
                    row_position += 1


    # If there are descriptions in other rows, they might need to be added in too.
    if descriptions_in_other_rows:
        descriptions_in_other_rows = list(filter(lambda i: i['Row'] not in data_rows and
                                                           i['Col'] not in columns_included,
                                                 descriptions_in_other_rows))
        for d in descriptions_in_other_rows:
            del d['Col']

        # for i in descriptions_in_other_rows:
        # i = {k:[elem for elem in v if elem is not np.nan] for k,v in i.items()}
    if descriptions_in_other_rows:
        spreadsheet_rows = list(range(min(data_rows), max(data_rows) + 1))
        correspondence = {}
        k = 0
        for i in spreadsheet_rows:
            if i in data_rows:
                correspondence[i] = k
                k += 1
        correspondence = pd.DataFrame(correspondence.items(), columns=['index', 'New_index'])
        #40379
        #last_stop = row_headings[0].count()
        descriptions_in_other_rows = pd.DataFrame(descriptions_in_other_rows)#.set_index('row_position')
        descriptions_in_other_rows = descriptions_in_other_rows.sort_values(by=['Row'])


        descriptions_in_other_rows.set_index('Row', inplace=True)
        descriptions_in_other_rows = descriptions_in_other_rows.reindex(range(max(data_rows)+1))
        descriptions_in_other_rows.ffill(axis=0, inplace=True)
        descriptions_in_other_rows['index'] = descriptions_in_other_rows.index
        descriptions_in_other_rows = descriptions_in_other_rows.merge(correspondence, on='index', how='left')

        descriptions_in_other_rows.drop(['row_position', 'index'], axis=1, inplace=True)
        #descriptions_in_other_rows.ffill(axis=0, inplace=True)
        #descriptions_in_other_rows = descriptions_in_other_rows[descriptions_in_other_rows.row_position >= 0]
        #descriptions_in_other_rows['row_position'] = descriptions_in_other_rows['row_position'].astype(int)
        descriptions_in_other_rows.set_index('New_index', inplace=True)
        descriptions_in_other_rows.rename_axis(None, inplace=True)
        row_headings = row_headings.join(descriptions_in_other_rows)

    return row_headings


def merged_data_subheadings_function(xl_workbook, sheet_name, merged_data_cols, data_cols, data_rows, rows,
                                          extra_rows,
                                          top_row, column_position=1):
    """ Identify additional merged subheadings that are in between data rows.
    Currently only works for one set of duplicate column names """

    first_row = min(data_rows)
    last_row = max(data_rows)

    sheet = xl_workbook.get_sheet_by_name(sheet_name)
    other_rows = set(i for i in range(last_row) if i not in data_rows and i > max(rows) - 1)
    all_rows = rows.union(other_rows)

    merged_meta_data_subheadings_potential = list(filter(lambda x: x[1] in other_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    merged_meta_data_cols = [el[0:3:2] for el in merged_meta_data_subheadings_potential]
    duplicates = list(set([ele for ele in merged_meta_data_cols
                           if merged_meta_data_cols.count(ele) > 1]))

    merged_meta_data_subheadings = []
    for i in merged_meta_data_subheadings_potential:
        for j in duplicates:
            if i[0] == j[0] and i[2] == j[1]:
                merged_meta_data_subheadings.append(i)

    duplicate_rows = set(el[1] for el in merged_meta_data_subheadings)

    first_col = min(data_cols)
    all_positions = []
    for i in all_rows:
        for j in data_cols:
            #all_positions.append((i, i + 1, j, j + 1))
            all_positions.append((j, i, j, i))

    all_merged_positions = []
    for i in merged_meta_data_subheadings:
        j = i[0]  # start column
        k = 0
        while j <= i[2]:
            for cells in range(i[0], i[0] + k + 1):
                #all_merged_positions.append((i[1], i[3], cells, i[0] + k))
                all_merged_positions.append((cells, i[1], i[0] + k, i[3]))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data_subheadings)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(1, 0))
    values = [0]
    values.extend(col for col in range(min(data_cols), max(data_cols)) if col not in data_cols)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_cols = dict(zip(keys, values))

    spreadsheet_rows = list(range(min(data_rows), max(data_rows) + 1))
    correspondence = {}
    k = 0
    for i in spreadsheet_rows:
        if i in data_rows:
            correspondence[i] = k
            k += 1
    correspondence = pd.DataFrame(correspondence.items(), columns=['index', 'New_index'])

    column_subheadings = pd.DataFrame()
    for j in duplicate_rows:
        for i in merged_meta_data_extended:
            if i[1] == j:
                #col = i[0]
                columns_in_df = i[0] - first_col + 1  # row position in df (in other words, the column number)
                rows_in_spreadsheet = i[1]
                # Filter out entries that occur in empty columns
                empty_cols_filtered = dict(filter(lambda elem: elem[1] <= i[3], empty_cols.items()))
                if empty_cols_filtered:
                    columns_in_df = columns_in_df - max(empty_cols_filtered, key=empty_cols_filtered.get)
                for k in range(i[0], i[2] + 1):
                    cell = sheet.cell(row=i[1], column=i[0])
                    column_subheadings.loc[rows_in_spreadsheet, "column_subheading"] = cell.value

        column_position += 1
    column_subheadings = column_subheadings.sort_index()
    column_subheadings = column_subheadings.append(pd.Series(name=last_row))

    df = pd.DataFrame({
        'column_heading': range(last_row + 1)})
    df = df.drop(['column_heading'], axis=1)
    df = df.join(column_subheadings)
    column_subheadings = df.ffill(axis=0).reset_index()

    column_subheadings = column_subheadings.merge(correspondence, on='index', how='right')
    column_subheadings = column_subheadings.drop(['index'], axis=1)
    column_subheadings = column_subheadings.rename(columns={"New_index": "index"})
    column_subheadings.set_index('index', inplace=True)
    column_subheadings = column_subheadings.rename_axis(None)

    return column_subheadings

