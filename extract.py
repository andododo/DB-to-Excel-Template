import os
import db_connection
import formatting
import tempfile

def title_cell(value, worksheet):
    cell = 'A1'
    worksheet[cell].value = value
    formatting.format_title(worksheet, cell)

def single_value(tbl_name, column_header, header_cell, value_cell, worksheet, header_mapping, application_id, request_id):
    conn = db_connection.connect()
    cursor = conn.cursor()
    query = f"""SELECT TOP 1 [{column_header}] 
            FROM {tbl_name} 
            WHERE [lngApplicationId]={application_id}
            AND [strRequestId]='{request_id}'
            """
    cursor.execute(query)
    value = cursor.fetchone()[0]

    header_name = header_mapping.get(column_header, column_header)
    worksheet[header_cell] = header_name
    formatting.format_upper(worksheet, header_cell)
    worksheet[value_cell] = value
    formatting.format_data(worksheet, value_cell)

    cursor.close()
    conn.close()

def column_values(tbl_name, column_header, header_cell, worksheet, header_mapping, application_id, request_id):
    conn = db_connection.connect()
    cursor = conn.cursor()

    query = f"""SELECT [{column_header}]
            FROM {tbl_name}
            WHERE [lngApplicationId]={application_id}
            AND [strRequestId]='{request_id}'
            AND [strRefName]='CAPA Report'
            """
    cursor.execute(query)
    values = [row[0] for row in cursor.fetchall()]

    header_name = header_mapping.get(column_header, column_header)
    worksheet[header_cell] = header_name
    formatting.format_lower(worksheet, header_cell)
    for i, value in enumerate(values, start=1):
        cell = f"{header_cell[0]}{int(header_cell[1:]) + i}"
        worksheet[cell] = value
        formatting.format_data(worksheet, cell)

    cursor.close()
    conn.close()

def save_workbook(workbook, temp_dir, temp_name): 
    # create a temporary directory
    # temp_dir = tempfile.mkdtemp()
    # print(temp_dir)

    # file path in the temporary directory
    output_file_name = temp_name + ".xlsx"
    output_file = os.path.join(temp_dir, output_file_name)

    # save the workbook
    workbook.save(output_file)

    # print(output_file)

    return output_file