# libraries
import openpyxl
import sys

# modules
import extract
import formatting
import header

def main(application_id, request_id, temp_dir, temp_name):
    '''
    CREATE EXCEL FILE
    =openpyxl functions
    '''
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    extract.title_cell('CAPA VALIDATION TRACKER', worksheet)

    '''
    -SINGLE DATA COLUMNS-
    PARAMETERS
    (<table name>, <header name, <header cell position>, <data cell position>)
    =worksheet is pass-through value for openpyxl library
    =mapping is for header names
    '''
    header_mapping_upper = header.get_header_upper()
    extract.single_value('tbl_v3_RequestDetails', 'strCol_14', 'A2', 'B2', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_15', 'D2', 'E2', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_16', 'G2', 'H2', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_17', 'J2', 'K2', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_18', 'A3', 'B3', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_19', 'D3', 'E3', worksheet, header_mapping_upper, application_id, request_id)
    extract.single_value('tbl_v3_RequestDetails', 'strCol_20', 'G3', 'H3', worksheet, header_mapping_upper, application_id, request_id)

    '''
    -MULTIPLE DATA COLUMNS-
    PARAMETERS
    (<table name>, <header name, <header cell position>)
    =column data will be placed under header cell position
    PERMANENT PARAMETERS
    =worksheet is pass-through value for openpyxl library
    =mapping is for header names
    '''
    header_mapping_lower = header.get_header_lower()
    extract.column_values('tbl_v3_Extends', 'strCol_08', 'A4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_09', 'B4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_10', 'C4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_11', 'E4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_12', 'D4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_13', 'F4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_15', 'G4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_16', 'H4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_17', 'I4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_18', 'J4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_19', 'K4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_20', 'L4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_21', 'M4', worksheet, header_mapping_lower, application_id, request_id)
    extract.column_values('tbl_v3_Extends', 'strCol_22', 'N4', worksheet, header_mapping_lower, application_id, request_id)

    '''
    -ADJUST COL WIDTH AND ROW HEIGHT-
    -MERGE NEEDED CELLS-
    '''
    formatting.merge_cells(workbook)
    formatting.adjust_dimensions(workbook)

    '''
    -SAVE EXCEL-
    PARAMETERS
    <file name>
    '''
    # file_name = 'capa-validation-tracker-temp'
    output_file = extract.save_workbook(workbook, temp_dir, temp_name)

    return output_file

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Usage: python initialize.py <application_id> <request_id> <temp_dir> <temp_name>")
        sys.exit(1)

    application_id = sys.argv[1]
    request_id = sys.argv[2]
    temp_dir = sys.argv[3]
    temp_name = sys.argv[4]

    output_file = main(application_id, request_id, temp_dir, temp_name)
    print(output_file)