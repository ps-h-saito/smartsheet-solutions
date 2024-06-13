# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
#############################################################################
###                                                                       ###
###  sheetdata write insert                                               ###
###    parameters:seach_sheet_name, process_workspace_id, insert_sheet_id ###
###                                                                       ###
###  (sample)                                                             ###
###   py No.01-1_sheetdata-write-insert-python.py "Sample Sheet" 2643982103668612 3143381366558596 ###
###                                                                       ###
#############################################################################
import azure.functions as func
import datetime
import json
import logging
import smartsheet
import os

app = func.FunctionApp()

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
newColumn_map = {}
insColumn_map = {}

# Setting Start --------------------------------

# Column List( [fromSheet_ColumnName,toSheet_ColumnName] Reflects only specified columns)
column_key = ["Name","Name"]
column_list = [["Status","Status"],["Remaining","Remaining"]]

# Setting End   --------------------------------

seach_sheet_name_Fixed = "Sample Sheet2"
seach_workspace_id_Fixed = 2643982103668612
insert_sheet_id_Fixed = 3143381366558596

# Helper function to find cell in a row
def get_cell_by_new_column_name(row, column_name):
    column_id = newColumn_map[column_name]
    return row.get_column(column_id)

# Get the sheet ID from the sheet name.
# If a workspace is specified, it is retrieved from the sheet in the workspace.
def get_sheet_name_from_id(sheet_name,smart,seach_sheet_name,seach_workspace_id,insert_sheet_id):
    logging.info(f"seach_sheet_name:{seach_sheet_name} seach_workspace_id:{seach_workspace_id} insert_sheet_id:{insert_sheet_id}")
    sheet_id = None
    response = smart.Sheets.list_sheets(include_all=True)
    for data in response.data:
        if data.name == sheet_name:
            sheet = smart.Sheets.get_sheet(data.id)
            if sheet.workspace is not None:
                if str(sheet.workspace.id) == str(seach_workspace_id):
                    sheet_id = data.id
                    break
            else:
                sheet_id = data.id
                break
    return sheet_id

# evaluate row and build insert data.
def evaluate_row_and_build_insert_data(source_row):
    row = smartsheet.models.Row()
    row.to_last = True
    name_cell = get_cell_by_new_column_name(source_row, column_key[0])
    if name_cell.value is not None:
        row.cells.append({
          'column_id': insColumn_map[column_key[1]],
          'value': name_cell.value or "",
          'strict': False
        })
        for item in column_list:
            cell = get_cell_by_new_column_name(source_row, item[0])
            row.cells.append({
              'column_id': insColumn_map[item[1]],
              'value': cell.value or "",
              'strict': False
            })
        return row
    else:
        return None


@app.route(route="sheetdata_write_insert", auth_level=func.AuthLevel.ANONYMOUS)
def sheetdata_write_insert(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（sheetdata_write_insert）")
    print("Start!!")

    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    try:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN_AZURE']
    except:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN']
    smart = smartsheet.Smartsheet(access_token)
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    # パラメータセット
    seach_sheet_name = req.params.get('seach_sheet_name')
    seach_workspace_id = req.params.get('seach_workspace_id')
    insert_sheet_id = req.params.get('insert_sheet_id')

    if seach_sheet_name is None:
        seach_sheet_name = seach_sheet_name_Fixed
        seach_workspace_id = seach_workspace_id_Fixed
        insert_sheet_id = insert_sheet_id_Fixed

    # Get workspace information from workspace ID
    seach_workspace_name = ""
    if seach_workspace_id is not None:
        workspace = smart.Workspaces.get_workspace(seach_workspace_id)
        if workspace is not None:
            seach_workspace_name = workspace.name

    # Get sheet information from sheet Name
    sheet_id = None
    if seach_sheet_name is not None:
        sheet_id = get_sheet_name_from_id(seach_sheet_name,smart,seach_sheet_name,seach_workspace_id,insert_sheet_id)

    rtc=""
    if sheet_id is not None:
        # Load entire sheet
        # Import Sheet
        newSheet = smart.Sheets.get_sheet(sheet_id)
        # Insert Sheet
        insSheet = smart.Sheets.get_sheet(insert_sheet_id)
    
        print("Loaded " + str(len(newSheet.rows)) + " rows ImportSheet name: " + newSheet.name + " sheetId: " + str(newSheet.id) + " InsertSheet name: " + insSheet.name + " sheetId: " + str(insSheet.id))

        # Build column map for later reference - translates column names to column id
        for column in newSheet.columns:
            newColumn_map[column.title] = column.id
        for column in insSheet.columns:
            insColumn_map[column.title] = column.id
    
        # Insert Rows Data
        insDataCnt = 0
        for row in newSheet.rows:
            # Insert Row Data Create
            insRow = evaluate_row_and_build_insert_data(row)        
            # Insert Row Data
            if insRow :
                insDataCnt += 1
                response = smart.Sheets.add_rows(
                    insert_sheet_id,
                    [insRow])

        # Import Sheet Move
        dt_now = datetime.datetime.now()
        after_sheet_name = newSheet.name + "_" + dt_now.strftime('%Y%m%d')
        updated_sheet = smart.Sheets.update_sheet(
        newSheet.id,
        smartsheet.models.Sheet({
            'name': after_sheet_name
        })
        )

        if seach_workspace_id is not None:
            rtc = "■対象のシートを取込み、データを追加しました。 取込ワークスペース指定:あり (id:" + str(seach_workspace_id) + " name:" + seach_workspace_name + ")"
        else:
            rtc = "■対象のシートを取込み、データを追加しました。 取込ワークスペース指定:無し"
        rtc += " ■処理情報[ 検索シート名:" + newSheet.name + " (id:" + str(newSheet.id) + ") 追加先シート名:" + insSheet.name + " (id:" + str(insSheet.id) + ") 追加データ数:" + str(insDataCnt) + " 取込対象変更後シート名:" + after_sheet_name + "]"
        return func.HttpResponse(rtc)
    else:
        if seach_workspace_id is not None:
            rtc = "■対象のシートは存在しません。 取込ワークスペース指定:あり (id:" + str(seach_workspace_id) + " name:" + seach_workspace_name + ")"
        else:
            rtc = "■対象のシートは存在しません。 取込ワークスペース指定:無し"
        return func.HttpResponse(rtc)
