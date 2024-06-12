# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
import azure.functions as func
import datetime
import json
import logging
import smartsheet
import os

app = func.FunctionApp()

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
mstColumn_map = {}
updColumn_map = {}

# Setting Start --------------------------------

master_sheet_id_Fixed = 8473541410246532
update_sheet_id_Fixed = 1860795012960132
master_column_name_Fixed = "小項目名"
update_column_name_Fixed = "項目名"

# Setting End   --------------------------------

# Helper function to find cell in a row
def get_cell_by_mst_column_name(row, column_name):
    column_id = mstColumn_map[column_name]
    return row.get_column(column_id)

def get_cell_by_upd_column_name(row, column_name):
    column_id = updColumn_map[column_name]
    return row.get_column(column_id)

# dropdpwnlist new option from master data
def evaluate_update_option_data(target_sheet, target_column_name):
  columnId = mstColumn_map[target_column_name]
  optionList = []
  for row in target_sheet.rows:
    for cell in row.cells:
      if cell.column_id == columnId :
        optionList.append(cell.value)

  return optionList


@app.timer_trigger(schedule="0 */10 * * * *", arg_name="myTimer", run_on_startup=False,
              use_monitor=False) 
def dropdownlist_update_timer(myTimer: func.TimerRequest) -> None:
    
    logging.info('Python Timer trigger function processed a request.')
    logging.info(f"▼処理を開始します（dropdownlist_update_timer）")

    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    try:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN_AZURE']
    except:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN']
    smart = smartsheet.Smartsheet(access_token)
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    # パラメータセット
    #  master Sheet Id
    master_sheet_id = None
    #  update Sheet Id
    update_sheet_id = None
    #  Master Target Column Name
    master_column_name = None
    #  Update Target Dropdown Column Name
    update_column_name = None

    # master_sheet_id = req.params.get('master_sheet_id')
    # update_sheet_id = req.params.get('update_sheet_id')
    # master_column_name = req.params.get('master_column_name')
    # update_column_name = req.params.get('update_column_name')

    if master_sheet_id is None:
        master_sheet_id = master_sheet_id_Fixed
        update_sheet_id = update_sheet_id_Fixed
        master_column_name = master_column_name_Fixed
        update_column_name = update_column_name_Fixed

    # Load entire sheet
    # master Sheet
    mstSheet = smart.Sheets.get_sheet(master_sheet_id)
    # update Sheet
    updSheet = smart.Sheets.get_sheet(update_sheet_id)

    logging.info("Loaded " + str(len(mstSheet.rows)) + " rows MasterSheet name: " + mstSheet.name + " sheetId: " + str(mstSheet.id) + " updateSheet name: " + updSheet.name + " sheetId: " + str(updSheet.id))

    # Build column map for later reference - translates column names to column id
    for column in mstSheet.columns:
        mstColumn_map[column.title] = column.id
    for column in updSheet.columns:
        updColumn_map[column.title] = column.id

    logging.info("Column Info update_sheet_id:" + str(update_sheet_id) + " update_column_id:" + str(updColumn_map[update_column_name]))

    # Specify column properties get
    column = smart.Sheets.get_column(
    update_sheet_id,       # sheet_id
    updColumn_map[update_column_name]       # column_id
    )

    column_spec = smart.models.Column({
    'title': column.title,
    'type': 'PICKLIST',
    'options': evaluate_update_option_data(mstSheet, master_column_name),
    'index': 2
    })

    # Update column
    response = smart.Sheets.update_column(
    update_sheet_id,       # sheet_id
    updColumn_map[update_column_name] ,       # column_id
    column_spec)
    updated_column = response.result

    logging.info("■ドロップダウンリストの更新が完了しました")
