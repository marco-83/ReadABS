#  Functions to locate tables in ABS spreadsheets and save location and descriptions to class TableData

from xlrd.sheet import ctype_text
import pandas as pd
import copy
import xlsxwriter
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
        #self.non_blank_rows_above_data = set()
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
            'Rows': str(min(self.rows)+1) + ':' + str(max(self.rows)+1),
            'Columns': xlsxwriter.utility.xl_col_to_name(min(self.cols)) + ':' +
                       xlsxwriter.utility.xl_col_to_name(max(self.cols)),
        }


def import_spreadsheet(excel_workbook="81670do001_201718.xls", filter_tabs=True):
    #xl_workbook = xlrd.open_workbook(excel_workbook, formatting_info=True)
    # TODO: Block macros
    xl_workbook = excel_workbook
    data_sheets = []
    if filter_tabs:
        for sheet in xl_workbook.sheets():
            if sheet.name.startswith("Table") or sheet.name.startswith("Data"):
                data_sheets.append(sheet.name)
    else:
        for sheet in xl_workbook.sheets():
            data_sheets.append(sheet.name)

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

            #print("first_col", first_col, "first_row", first_row, "table.sheet_name", table.sheet_name)
            # Find row descriptions
            table.row_descriptions, table.table_type = locate_row_descriptions(xl_workbook, first_col, first_row,
                                                                               table.sheet_name, table.cols)

            if table.row_descriptions == "Could not locate row headings":
                continue

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
                                                                                 table.top_header_row,
                                                                                 table.last_row_in_sheet)
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
    sheet = xl_workbook.sheet_by_name(sheet_name)
    for r in range(sheet.nrows):
        row = sheet.row(r)
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
        sheet = xl_workbook.sheet_by_name(s)
        found_table = False
        r = 0
        found_data = False
        looking_for_multiple_tables = False
        blank_row = False
        blank_row_count = 0
        quit_loop = False
        date_cols = []

        if data_type == "Time series":
            for i in range(sheet.nrows):
                row = sheet.row(i)
                if quit_loop:
                    break
                for idx, cell_obj in enumerate(row):
                    if cell_obj.value == "Series ID":
                        start_row = i + 1
                        quit_loop = True
                        break
        else:
            start_row = 0

        quit_loop = False

        if 'start_row' not in locals():
            start_row = 0

        # Find data
        for i in range(start_row, sheet.nrows):
            # if i < 9:
            #     continue
            if quit_loop:
                break
            row = sheet.row(i)
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
                            date_cols.append(idx)
                            continue
                # Return the cell 'weight' (700 is bold, 400 is 'normal')
                rd_xf = xl_workbook.xf_list[sheet.cell_xf_index(i, idx)]
                cell_font = xl_workbook.font_list[rd_xf.font_index].weight
                #cell_format = xl_workbook.format_map[rd_xf.format_key].format_str
                cell_type = ctype_text.get(cell_obj.ctype, 'unknown type')
                # if idx in [0,4]:
                #     print('row', i, 'column', idx, 'cell_type', cell_type, 'cell_obj.value', cell_obj.value, 'cell_format',
                #           cell_format)
                # Check if it is a total row (p.s. this is very specific, sometimes total row is at the top and bold
                if cell_type == "number" and not found_data and idx != 0:  # Top left data cell
                    left_of_data = sheet.row(i)[idx - 1].value
                #if cell_type == "number" and cell_format != "General" and \
                if cell_type == "number" and (cell_font != 700 or left_of_data == "Total") and idx not in date_cols:
                    # Store info on location of data column
                    found_data = True
                    if not found_table:
                        tables.append(TableData(sheet_name=sheet.name))
                        table_number += 1
                        found_table = True
                    tables[table_number].add_row(i)
                    tables[table_number].add_col(idx)
                    c += 1
                # Sometimes there is a 'total' row that is bold. Let's not skip it.
                # If the table has been started, allow bold rows.
                # p.s this might be redundant if the total row is always called 'Total' as per above if statement
                #if found_table and cell_type == "number" and cell_format != "General" and cell_font == 700:
                if found_table and cell_type == "number" and cell_font == 700:
                    found_data = True
                    tables[table_number].add_row(i)
                    tables[table_number].add_col(idx)
                    c += 1
                if cell_type in ["number", "text"] and looking_for_multiple_tables is True:
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

    sheet = xl_workbook.sheet_by_name(sheet_name)
    row_descriptions = []
    # Where is the descriptor column?
    if table_type == "wide format":
        for c in range(0, first_col):
            cell = sheet.row(first_row)[c]
            if ctype_text.get(cell.ctype, 'unknown type') not in ["text", "xldate"]:
                continue
            else:
                row_descriptions.append(c)
                #break
    elif table_type == "long format":
        for c in range(0, first_col):
            cell = sheet.row(first_row-1)[c]
            if ctype_text.get(cell.ctype, 'unknown type') not in ["text", "xldate"]:
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

    sheet = xl_workbook.sheet_by_name(sheet_name)
    first_row = min(data_rows)
    last_row = max(data_rows)
    top_cell_row = min(data_rows)
    first_col = min(data_cols)

    # Identify the row(s) above the data that describe the columns.
    # Starts 1 row above the data, working up. If a blank row is found, then stop (unless no column headers have been \
    # identified yet).
    rows_above_data = range(0, first_row)

    non_blank_rows_above_data = set()
    blank_row = False
    found_a_non_blank_row = False
    for r in reversed(rows_above_data):
        if blank_row and found_a_non_blank_row:
            break
        row = sheet.row(r)
        for idx, cell_obj in enumerate(row):
            if idx >= first_col:
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if cell_obj.value == "":
                    blank_row = True
                else:
                    blank_row = False
                    found_a_non_blank_row = True
                    non_blank_rows_above_data.add(r)
                    break
    top_header_row = min(non_blank_rows_above_data)

    # How many levels of indentation are there, if any?
    # indentation_levels = set()
    indentation_levels = {}
    start_row = max(non_blank_rows_above_data)

    # Start 2 row above first row of data
    for i, c in enumerate(row_descriptions):
        indentation_levels[c] = set()
        for r in range(start_row, last_row + 1):
            indentation = xl_workbook.xf_list[sheet.cell_xf_index(r, row_descriptions[i])].alignment.indent_level
            if sheet.row(r)[row_descriptions[i]].value != '':
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
                indentation_cell = xl_workbook.xf_list[sheet.cell_xf_index(cell_row, columns_with_indentation[idx])].alignment.indent_level
                while indentation_cell > i:
                    cell_row -= 1
                    indentation_cell = xl_workbook.xf_list[sheet.cell_xf_index(cell_row, columns_with_indentation[idx])].alignment.indent_level
                    # print('indentation_cell', indentation_cell)
                    if cell_row < top_cell_row:
                        top_cell_row = cell_row
            # column += 1
    # if sheet_name == "Table 1":
    #     print('top_cell_row_indentation', top_cell_row)
    return indentation_levels, columns_with_indentation, top_cell_row, top_header_row

    # # Only run if there is 1 info col
    # # TODO: Now redundant: remove if statement
    # if len(row_descriptions) == 1:
    #     # How many levels of indentation are there, if any?
    #     indentation_levels = set()
    #     start_row = max(non_blank_rows_above_data)
    #     # Start 2 row above first row of data
    #     for r in range(start_row, last_row):
    #         indentation = xl_workbook.xf_list[sheet.cell_xf_index(r, row_descriptions[0])].alignment.indent_level
    #         indentation_levels.add(indentation)
    #
    #     indentation_levels = list(indentation_levels)
    #     indentation_levels.sort()
    #
    #     for r in data_rows:
    #         column = 0
    #         for i in indentation_levels:
    #             cell_row = r
    #             indentation_cell = xl_workbook.xf_list[sheet.cell_xf_index(cell_row, row_descriptions[0])].alignment.indent_level
    #             while indentation_cell > i:
    #                 cell_row -= 1
    #                 indentation_cell = xl_workbook.xf_list[sheet.cell_xf_index(cell_row, row_descriptions[0])].alignment.indent_level
    #                 if cell_row < top_cell_row:
    #                     top_cell_row = cell_row
    #             column += 1
    #     return indentation_levels, top_cell_row, top_header_row
    # else:
    #     return [], top_cell_row, top_header_row


def describe_row_headers(data_rows, data_cols, sheet_name, row_descriptions, columns_with_indentation, xl_workbook):
    """ Function to find row descriptions. Only run if there is more than 1 row descriptor column,
    and check for merged cells"""

    sheet = xl_workbook.sheet_by_name(sheet_name)
    first_row = min(data_rows)
    top_cell_row = min(data_rows)
    first_col = min(data_cols)
    other_columns = set(i for i in row_descriptions if i not in columns_with_indentation)

    # Identify the row(s) above the data that describe the columns.
    # Starts 1 row above the data, working up. If a blank row is found, then stop (unless no column headers have been \
    # identified yet).
    rows_above_data = range(0, first_row)

    non_blank_rows_above_data = set()
    blank_row = False
    found_a_non_blank_row = False
    for r in reversed(rows_above_data):
        if blank_row and found_a_non_blank_row:
            break
        row = sheet.row(r)
        for idx, cell_obj in enumerate(row):
            if idx >= first_col:
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if cell_obj.value == "":
                    blank_row = True
                else:
                    blank_row = False
                    found_a_non_blank_row = True
                    non_blank_rows_above_data.add(r)
                    break
    #top_header_row = min(non_blank_rows_above_data)

    all_mergers = sheet.merged_cells

    merged_meta_data_row_headings = []
    merged_meta_data_col_headings = []
    for i in all_mergers:
        # Only keep the merged cells that are above the data; not to the left and not to the right
        if i[0] in data_rows and i[2] < first_col:
            merged_meta_data_row_headings.append(i)
        elif i[2] < first_col:
            merged_meta_data_col_headings.append(i)

    merged_meta_data_col_headings = [a_tuple[0] for a_tuple in merged_meta_data_col_headings]

    # if sheet_name == "Table 4":
    #     print("all_mergers", all_mergers, "merged_meta_data_col_headings", merged_meta_data_col_headings)

    # Look for a header row above the row descriptions.
    # The header row must have something in each cell above the row descriptions, otherwise return None.
    row_descriptions_header_row = None
    for r in reversed(range(1, min(data_rows))):  # Note: 'range' does not include end point
        if r not in merged_meta_data_col_headings and r > top_cell_row - 2:  # 2 above data is a bit arbitrary
            if row_descriptions_header_row:
                break
            for c in other_columns:
                if not sheet.row(r)[c].value:
                    row_descriptions_header_row = None
                    continue
                else:
                    row_descriptions_header_row = r

    row_titles = {}
    if row_descriptions_header_row:
        for i, c in enumerate(other_columns):
            title = sheet.row(row_descriptions_header_row)[c].value
            if title and title != '':
                row_titles[c] = sheet.row(row_descriptions_header_row)[c].value
            else:
                row_titles[c] = "Row_description_title_sub_" + str(i)

    else:
        for i, c in enumerate(other_columns):
            row_titles[c] = "Row_description_title_sub_" + str(i)
    # if sheet_name == "Table 1":
    #     print("merged_meta_data_row_headings", merged_meta_data_row_headings)
    return merged_meta_data_row_headings, row_descriptions_header_row, row_titles


def describe_col_headings_timeseries(xl_workbook, sheet_name, data_rows, data_cols, top_header_row, last_row_in_sheet):

    first_row = min(data_rows)
    first_col = min(data_cols)
    sheet = xl_workbook.sheet_by_name(sheet_name)

    rows_above_data = range(0, first_row)
    quit_loop = False
    for i in range(last_row_in_sheet+1):
        row = sheet.row(i)
        if quit_loop:
            break
        for idx, cell_obj in enumerate(row):
            if cell_obj.value == "Series ID":
                series_id_position = [row, idx]  # location of "Series ID" field
                quit_loop = True
                break

    if 'series_id_position' not in locals():
        series_id_position = [9, 0]

    column_header_locations = set()
    for r in range(top_header_row, first_row):
        if sheet.cell_value(rowx=r, colx=first_col) is not None:
            column_header_locations.add(r)
    column_titles = {}
    for r in column_header_locations:
        column_titles[r] = sheet.cell_value(rowx=r, colx=series_id_position[1])
        if column_titles[r] == '':
            column_titles[r] = "Description"

    return column_titles, column_header_locations


def describe_col_headings(xl_workbook, sheet_name, data_rows, data_cols, row_descriptions, top_row, top_header_row,
                          row_descriptions_header_row, spreadsheet_type, last_row_in_sheet):
    """ Find column headings. There might be multiple column headings (above each other) that might be units ($, %, etc)
    or they might be merged cells """

    first_row = min(data_rows)
    last_col = max(data_cols)
    first_col = min(data_cols)
    sheet = xl_workbook.sheet_by_name(sheet_name)

    # Find merged column headings
    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d is the bottom right coordinate (row / col, starting with 1) where the merge finishes (who knows why?)
    assert isinstance(sheet.merged_cells, object)
    all_mergers = sheet.merged_cells

    rows_above_data = range(0, first_row)

    column_header_locations = set()
    blank_row = False
    found_a_non_blank_row = False
    found_a_column_header = False
    #print("all_mergers", all_mergers)

    def check_rows(idx, merged_meta_data):
        if merged_meta_data:
            merged_col1 = list(zip(*merged_meta_data))[2]
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
        row = sheet.row(r)
        merged_meta_data = list(filter(lambda x: x[0] == r, all_mergers))
        for idx, cell_obj in enumerate(row):
            # Return the cell 'weight' (700 is bold, 400 is 'normal')
            rd_xf = xl_workbook.xf_list[sheet.cell_xf_index(r, idx)]
            cell_font = xl_workbook.font_list[rd_xf.font_index].weight
            cell_type = ctype_text.get(cell_obj.ctype, 'unknown type')
            # If the column index is greater than the first data column or is in the first column of merged cells
            if check_rows(idx, merged_meta_data):
                # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                if cell_obj.value == "":
                    blank_row = True
                # At least one of the column headers must be text or bold for this to work
                elif cell_type == "text" or cell_font == 700:
                    blank_row = False
                    found_a_non_blank_row = True
                    found_a_column_header = True
                    column_header_locations.add(r)
                    break
                else:
                    blank_row = False
                    found_a_non_blank_row = True

    #column_header_locations.discard(top_row)

    # Extra meta data above top left cell
    extra_meta_data = set()
    if spreadsheet_type == "Data cube":
        if row_descriptions_header_row:
            column_headers_already_included = {top_row, row_descriptions_header_row}
        else:
            column_headers_already_included = {top_row}
        if all_mergers:
            for r in column_header_locations:
                mergers_filtered = [tup for tup in all_mergers if tup[0] == r]
                if mergers_filtered:
                    for c in row_descriptions:
                        if c in list(zip(*mergers_filtered))[2]:
                            column_headers_already_included.add(r)
                            # if sheet_name == "Table 1":
                            #     print('r', r, 'c', c, 'all_mergers', all_mergers, 'mergers_filtered', mergers_filtered)

        rows_above_data = list(filter(lambda i: i not in column_headers_already_included, [*rows_above_data]))

        blank_row = False
        for r in reversed(rows_above_data):
            if blank_row:
                break
            row = sheet.row(r)
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx in row_descriptions:
                    rd_xf = xl_workbook.xf_list[sheet.cell_xf_index(r, idx)]
                    cell_format = xl_workbook.format_map[rd_xf.format_key].format_str
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "":
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif ctype_text.get(cell_obj.ctype, 'unknown type') == "text" or cell_format == 'General':
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False

        blank_row = False
        for r in rows_above_data:
            if blank_row:
                break
            row = sheet.row(r)
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx in row_descriptions:
                    rd_xf = xl_workbook.xf_list[sheet.cell_xf_index(r, idx)]
                    cell_format = xl_workbook.format_map[rd_xf.format_key].format_str
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "":
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif ctype_text.get(cell_obj.ctype, 'unknown type') == "text" or cell_format == 'General':
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False
    # if sheet_name == "Table 1":
    #     print('extra_meta_data', extra_meta_data, 'row_descriptions_header_row', row_descriptions_header_row,
    #           'top_row', top_row, 'top_header_row', top_header_row, 'rows_above_data', rows_above_data,
    #           'column_headers_already_included', column_headers_already_included, 'row_descriptions', row_descriptions)
    else:
        columns_to_evaluate = range(0, max(row_descriptions))
        rows_above_top_header_row = range(0, top_header_row)
        column_headers_already_included = {top_header_row}
        if all_mergers:
            for r in column_header_locations:
                mergers_filtered = [tup for tup in all_mergers if tup[0] == r]
                if mergers_filtered:
                    for c in columns_to_evaluate:
                        if c in list(zip(*mergers_filtered))[2]:
                            column_headers_already_included.add(r)
                # for c in columns_to_evaluate:
                #     if r in list(zip(*all_mergers))[0] and c in list(zip(*all_mergers))[2]:
                #         column_headers_already_included.add(r)

        blank_row = False
        rows_above_data = list(filter(lambda i: i not in column_headers_already_included, [*rows_above_top_header_row]))

        for r in reversed(rows_above_data):
            row = sheet.row(r)
            for idx, cell_obj in enumerate(row):
                # If the column index is greater than the first data column or is in the first column of merged cells
                if idx in columns_to_evaluate:
                    # As soon as something is found in a cell, it is not a blank row and break the loop, go to next row
                    if cell_obj.value == "":
                        blank_row = True
                    # At least one of the column headers must be text for this to work
                    elif ctype_text.get(cell_obj.ctype, 'unknown type') == "text":
                        blank_row = False
                        extra_meta_data.add(r)
                        break
                    else:
                        blank_row = False


    other_rows = set(i for i in range(last_row_in_sheet+1) if i not in data_rows and i > max(column_header_locations))
    all_rows = column_header_locations.union(other_rows)
    merged_meta_data = []
    for i in all_mergers:
        # Only keep the merged cells that are above the data; not to the left and not to the right
        #if i[0] in column_headers_and_rows_above_data and \
        #        i[2] >= first_col and i[1] - 1 in column_headers_and_rows_above_data and i[3] - 1 <= last_col:
        if i[0] in all_rows and i[2] <= last_col:
            merged_meta_data.append(i)

    return merged_meta_data, column_header_locations, extra_meta_data


def merged_data_function(xl_workbook, sheet_name, merged_data_cols, data_cols, data_rows,
                         extra_rows, last_row_in_sheet, spreadsheet_type, column_header_locations, column_position=1):
    """ Function to extract data from merged cells
    merged_data_cols is a list of tuples. Each tuple is in the format used by xlrd function merged_cells """

    sheet = xl_workbook.sheet_by_name(sheet_name)
    last_row = max(data_rows)
    other_rows = set(i for i in range(last_row_in_sheet+1) if i not in data_rows and i > max(column_header_locations))
    all_rows = column_header_locations.union(other_rows)

    column_headings = pd.DataFrame()
    merged_meta_data = list(filter(lambda x: x[0] in all_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    other_rows = set(i for i in range(last_row) if i not in data_rows and i > max(column_header_locations) - 1)
    all_rows = column_header_locations.union(other_rows)

    merged_meta_data_subheadings_potential = list(filter(lambda x: x[0] in other_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    merged_meta_data_last_2_elements = [el[2:] for el in merged_meta_data_subheadings_potential]
    duplicates = list(set([ele for ele in merged_meta_data_last_2_elements
                           if merged_meta_data_last_2_elements.count(ele) > 1]))

    merged_meta_data_subheadings = []
    for i in merged_meta_data_subheadings_potential:
        for j in duplicates:
            if i[2] == j[0] and i[3] == j[1]:
                merged_meta_data_subheadings.append(i)

    # duplicate_rows = set(el[0] for el in merged_meta_data_subheadings)
    #
    # merged_meta_data_subheadings = []
    # for i in merged_meta_data:
    #     for j in duplicates:
    #         if i[2] == j[0] and i[3] == j[1]:
    #             merged_meta_data_subheadings.append(i)

    # Remove the subheading rows
    merged_meta_data = [x for x in merged_meta_data if x not in merged_meta_data_subheadings]
    subheading_rows = [i[0] for i in merged_meta_data_subheadings]
    rows_not_subheadings = [x for x in column_header_locations if x not in subheading_rows]

    # Find how the merged data relates to the columns
    first_col = min(data_cols)
    all_positions = []
    for i in rows_not_subheadings:
        for j in data_cols:
            all_positions.append((i, i + 1, j, j + 1))

    all_merged_positions = []
    for i in merged_meta_data:
        j = i[2]  # start column
        k = 1
        while j < i[3]:
            for cells in range(i[2], i[2] + k):
                all_merged_positions.append((i[0], i[1], cells, i[2] + k))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(0, 2))

    values = [0]
    values.extend(col for col in range(min(data_cols), max(data_cols)) if col not in data_cols)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_cols = dict(zip(keys, values))

    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d is the bottom right coordinate (row / col, starting with 1) where the merge finishes (who knows why?)
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
            if i[0] == r:
                row_position = i[2] - first_col  # row position in df (in other words, the column number)
                # Filter out entries that occur in empty columns
                empty_cols_filtered = dict(filter(lambda elem: elem[1] < i[3], empty_cols.items()))
                if empty_cols_filtered:
                    row_position = row_position - max(empty_cols_filtered, key=empty_cols_filtered.get)
                for k in range(i[2], i[3]):
                    cell = sheet.row(i[0])[i[2]]
                    if row_position >= 0:
                        if ctype_text.get(cell.ctype, 'unknown type') == "xldate":
                            column_headings.loc[row_position, column_heading] = pd.to_datetime((cell.value - 25569) *
                                                                                               86400.0, unit='s').\
                                strftime('%d/%m/%Y')
                        else:
                            column_headings.loc[row_position, column_heading] = cell.value
                    row_position += 1

    columns_to_evaluate = [0] # Assume the extra row info is all in column A
    for j in extra_rows:
        column_heading = 'Col_desc_' + str(column_position)
        for i in columns_to_evaluate:
            cell = sheet.row(j)[i]
            for row_position in range(0, len(data_cols)):
                if ctype_text.get(cell.ctype, 'unknown type') == "xldate":
                    column_headings.loc[row_position, column_heading] = pd.to_datetime((cell.value - 25569) *
                                                                                       86400.0, unit='s'). \
                        strftime('%d/%m/%Y')
                else:
                    column_headings.loc[row_position, column_heading] = cell.value
        column_position += 1

    return column_headings


def merged_data_row_headings_function(xl_workbook, sheet_name, merged_data_rows, data_rows, row_descriptions,
                                      row_titles, top_header_row):
    """ Function to extract data from merged cells
    merged_data_rows is a list of tuples. Each tuple is in the format used by xlrd function merged_cells """

    sheet = xl_workbook.sheet_by_name(sheet_name)

    other_rows = set(i for i in range(top_header_row + 1, max(data_rows)) if i not in data_rows)
    all_rows = data_rows.union(other_rows)

    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d is the bottom right coordinate (row / col, starting with 1) where the merge finishes

    row_headings = pd.DataFrame()
    all_positions = []
    for i in all_rows:
        for j in row_descriptions:
            all_positions.append((i, i + 1, j, j + 1))

    merged_meta_data_row_headings = list(filter(lambda x: x[0] in data_rows, merged_data_rows))
    if sheet_name == "Table 1":
        print("data_rows", data_rows)

    all_merged_positions = []
    for i in merged_meta_data_row_headings:
        j = i[0]  # start row
        k = 1
        while j < i[1]:
            for cells in range(i[0], i[0] + k):
                all_merged_positions.append((cells, i[0] + k, i[2], i[3]))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data_row_headings)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(0, 2))

    values = [0]
    values.extend(row for row in range(min(data_rows), max(data_rows)) if row not in data_rows)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_rows = dict(zip(keys, values))
    if sheet_name == "Table 1":
        print("empty_rows", empty_rows, "merged_meta_data_extended", merged_meta_data_extended)

    descriptions_in_other_rows = []
    columns_included = set()
    first_data_row = min(data_rows)
    for c in row_descriptions:
        column_heading = row_titles[c]
        for i in merged_meta_data_extended:
            if i[2] == c:
                row_position = i[0] - first_data_row  # row position in df (in other words, the row number)
                # Filter out entries that occur in empty columns
                empty_rows_filtered = dict(filter(lambda elem: elem[1] < i[1], empty_rows.items()))
                if empty_rows_filtered:
                    row_position = row_position - max(empty_rows_filtered, key=empty_rows_filtered.get)
                    for k in range(i[0], i[1]):
                        cell_value = sheet.row(i[0])[i[2]].value
                        if cell_value != '':
                            descriptions_in_other_rows.append({'Row': i[1], 'Col': i[2],
                                                               'row_position': row_position,
                                                               'Desc_row'+str(i[2]): cell_value})
                for k in range(i[0], i[1]):
                    cell = sheet.row(i[0])[i[2]]
                    if row_position >= 0 and i[0] in data_rows:
                        if ctype_text.get(cell.ctype, 'unknown type') == "xldate":
                            row_headings.loc[row_position, column_heading] = pd.to_datetime((cell.value - 25569) * 86400.0,
                                                                             unit='s').strftime('%d/%m/%Y')
                        else:
                            row_headings.loc[row_position, column_heading] = cell.value
                        if cell.value != '' and i[0] in data_rows:
                            columns_included.add(i[2])
                    row_position += 1

    if descriptions_in_other_rows:
        descriptions_in_other_rows = list(filter(lambda i: i['Row'] not in data_rows and
                                                           i['Col'] not in columns_included,
                                                 descriptions_in_other_rows))
        for d in descriptions_in_other_rows:
            del d['Col']

    if descriptions_in_other_rows:
        spreadsheet_rows = list(range(min(data_rows), max(data_rows) + 1))
        correspondence = {}
        k = 0
        for i in spreadsheet_rows:
            if i in data_rows:
                correspondence[i] = k
                k += 1
        correspondence = pd.DataFrame(correspondence.items(), columns=['index', 'New_index'])

        descriptions_in_other_rows = pd.DataFrame(descriptions_in_other_rows)
        descriptions_in_other_rows = descriptions_in_other_rows.sort_values(by=['Row'])

        descriptions_in_other_rows.set_index('Row', inplace=True)
        descriptions_in_other_rows = descriptions_in_other_rows.reindex(range(max(data_rows)+1))
        descriptions_in_other_rows.ffill(axis=0, inplace=True)
        descriptions_in_other_rows['index'] = descriptions_in_other_rows.index
        descriptions_in_other_rows = descriptions_in_other_rows.merge(correspondence, on='index', how='left')

        descriptions_in_other_rows.drop(['row_position', 'index'], axis=1, inplace=True)

        descriptions_in_other_rows.set_index('New_index', inplace=True)
        descriptions_in_other_rows.rename_axis(None, inplace=True)
        row_headings = row_headings.join(descriptions_in_other_rows)

    return row_headings


def merged_data_subheadings_function(xl_workbook, sheet_name, merged_data_cols, data_cols, data_rows, rows, extra_rows,
                                     top_row, column_position=1):
    """ Identify additional merged subheadings that are in between data rows.
    Currently only works for one set of duplicate column names """

    first_row = min(data_rows)
    last_row = max(data_rows)

    sheet = xl_workbook.sheet_by_name(sheet_name)
    other_rows = set(i for i in range(last_row) if i not in data_rows and i > max(rows)-1)
    all_rows = rows.union(other_rows)
    #print("other_rows", other_rows)
    merged_meta_data_subheadings_potential = list(filter(lambda x: x[0] in other_rows, merged_data_cols))

    # Get the merged items that have the same column dimensions. These are understood to be subheadings.
    merged_meta_data_last_2_elements = [el[2:] for el in merged_meta_data_subheadings_potential]
    duplicates = list(set([ele for ele in merged_meta_data_last_2_elements
                           if merged_meta_data_last_2_elements.count(ele) > 1]))

    merged_meta_data_subheadings = []
    for i in merged_meta_data_subheadings_potential:
        for j in duplicates:
            if i[2] == j[0] and i[3] == j[1]:
                merged_meta_data_subheadings.append(i)

    duplicate_rows = set(el[0] for el in merged_meta_data_subheadings)

    # print("merged_meta_data_subheadings", merged_meta_data_subheadings)
    # print("table.top_row", top_row)

    first_col = min(data_cols)
    all_positions = []
    for i in all_rows:
        for j in data_cols:
            all_positions.append((i, i + 1, j, j + 1))

    all_merged_positions = []
    for i in merged_meta_data_subheadings:
        j = i[2]  # start column
        k = 1
        while j < i[3]:
            for cells in range(i[2], i[2] + k):
                all_merged_positions.append((i[0], i[1], cells, i[2] + k))
            j += 1
            k += 1

    merged_meta_data_extended = copy.copy(merged_meta_data_subheadings)
    merged_meta_data_extended.extend(i for i in all_positions if i not in all_merged_positions)
    # Needs to be sorted to ensure the descriptions line up properly with the data
    merged_meta_data_extended.sort(key=itemgetter(0, 2))

    values = [0]
    values.extend(col for col in range(min(data_cols), max(data_cols)) if col not in data_cols)
    keys = [0]
    k = 1
    for v in values:
        keys.append(k)
        k += 1

    empty_cols = dict(zip(keys, values))

    # sheet.merged_cells returns a list of tuples. Each tuple has 4 elements a,b,c,d
    # a,c is the top-left coordinate (row / col, starting with 0) where the merge starts.
    # b,d is the bottom right coordinate (row / col, starting with 1) where the merge finishes (who knows why?)

    spreadsheet_rows = list(range(min(data_rows), max(data_rows)+1))
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
            if i[0] == j:
                col = i[2]
                columns_in_df = i[2] - first_col  # row position in df (in other words, the column number)
                rows_in_spreadsheet = i[0]
                # Filter out entries that occur in empty columns
                empty_cols_filtered = dict(filter(lambda elem: elem[1] < i[3], empty_cols.items()))
                if empty_cols_filtered:
                    columns_in_df = columns_in_df - max(empty_cols_filtered, key=empty_cols_filtered.get)
                for k in range(i[2], i[3]):
                    cell = sheet.row(i[0])[i[2]]
                    column_subheadings.loc[rows_in_spreadsheet, "column_subheading"] = cell.value

        column_position += 1
    column_subheadings = column_subheadings.sort_index()
    column_subheadings = column_subheadings.append(pd.Series(name=last_row))

    df = pd.DataFrame({
        'column_heading': range(last_row+1)})
    df = df.drop(['column_heading'], axis=1)
    df = df.join(column_subheadings)
    column_subheadings = df.ffill(axis=0).reset_index()

    column_subheadings = column_subheadings.merge(correspondence, on='index', how='right')
    column_subheadings = column_subheadings.drop(['index'], axis=1)
    column_subheadings = column_subheadings.rename(columns={"New_index": "index"})
    column_subheadings.set_index('index', inplace=True)
    column_subheadings = column_subheadings.rename_axis(None)

    # if sheet_name == "Table_1":
    #     print(column_subheadings)
    #     print(data_rows)
    #     print(spreadsheet_rows)

    return column_subheadings

