# Install the smartsheet sdk with the command: pip install smartsheet-python-sdk
#############################################################################
###                                                                       ###
###  smartsheet_azure_function                                            ###
###                                                                       ###
#############################################################################
import azure.functions as func
import datetime
import json
import logging
import smartsheet
import os
import time

app = func.FunctionApp()

# The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
newColumn_map = {}
insColumn_map = {}
mstColumn_map = {}
updColumn_map = {}
funcColumn_map = {}

# Setting Start --------------------------------

# Column List( [fromSheet_ColumnName,toSheet_ColumnName] Reflects only specified columns)
column_key = ["Name","Name"]
column_list = [["Status","Status"],["Remaining","Remaining"]]

# Setting End   --------------------------------

# defult initialize Start --------------------------------
seach_sheet_name_Fixed = "Sample Sheet"
seach_workspace_id_Fixed = 2643982103668612
insert_sheet_id_Fixed = 3143381366558596

master_sheet_id_Fixed = 8473541410246532
update_sheet_id_Fixed = 1860795012960132
master_column_name_Fixed = "小項目名"
update_column_name_Fixed = "項目名"
# defult initialize End   --------------------------------


# Helper function to find cell in a row
def get_cell_by_new_column_name(row, column_name):
    column_id = newColumn_map[column_name]
    return row.get_column(column_id)

def get_cell_by_mst_column_name(row, column_name):
    column_id = mstColumn_map[column_name]
    return row.get_column(column_id)

def get_cell_by_upd_column_name(row, column_name):
    column_id = updColumn_map[column_name]
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

# dropdpwnlist new option from master data
def evaluate_update_option_data(target_sheet, target_column_name):
  columnId = mstColumn_map[target_column_name]
  optionList = []
  for row in target_sheet.rows:
    for cell in row.cells:
      if cell.column_id == columnId :
        optionList.append(cell.value)

  return optionList


#############################################################################
###                                                                       ###
###  sheetdata write insert                                               ###
###    parameters:seach_sheet_name, seach_workspace_id, insert_sheet_id   ###
###                                                                       ###
#############################################################################
@app.route(route="sheetdata_write_insert", auth_level=func.AuthLevel.ANONYMOUS)
def sheetdata_write_insert(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（sheetdata_write_insert）")
    print("Start!!")

    # challenge check
    logging.info('challenge check')
    request_json=None
    try:
        request_json = req.get_json()
        logging.info(f'get_req:{request_json}')
    except ValueError:
        pass
    if request_json and "challenge" in request_json:
        logging.info('challenge response')
        return json.dumps({
            "smartsheetHookResponse": request_json['challenge']
        })
    logging.info('not challenge response')
    
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


#############################################################################
###                                                                       ###
###  sheetdata write insert timer                                         ###
###    parameters:seach_sheet_name, seach_workspace_id, insert_sheet_id   ###
###                                                                       ###
#############################################################################
@app.timer_trigger(schedule="0 0 6 * * *", arg_name="myTimer", run_on_startup=False,
              use_monitor=False) 
def sheetdata_write_insert_timer(myTimer: func.TimerRequest) -> None:
    
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（sheetdata_write_insert_timer）")
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
    seach_sheet_name = None
    seach_workspace_id = None
    insert_sheet_id = None

    # seach_sheet_name = myTimer.params.get('seach_sheet_name')
    # seach_workspace_id = myTimer.params.get('seach_workspace_id')
    # insert_sheet_id = myTimer.params.get('insert_sheet_id')

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
        logging.info(rtc)
    else:
        if seach_workspace_id is not None:
            rtc = "■対象のシートは存在しません。 取込ワークスペース指定:あり (id:" + str(seach_workspace_id) + " name:" + seach_workspace_name + ")"
        else:
            rtc = "■対象のシートは存在しません。 取込ワークスペース指定:無し"
        logging.info(rtc)


#############################################################################
###                                                                       ###
###  dropdownlist update                                                  ###
###    parameters:master_sheet_id, update_sheet_id, master_column_name, update_column_name ###
###                                                                       ###
#############################################################################
@app.route(route="dropdownlist_update", auth_level=func.AuthLevel.ANONYMOUS)
def dropdownlist_update(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（dropdownlist_update）")
    print("Start!!")

    # challenge check
    logging.info('challenge check')
    request_json=None
    try:
        request_json = req.get_json()
        logging.info(f'get_req:{request_json}')
    except ValueError:
        pass
    if request_json and "challenge" in request_json:
        logging.info('challenge response')
        return json.dumps({
            "smartsheetHookResponse": request_json['challenge']
        })
    logging.info('not challenge response')

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
    master_sheet_id = req.params.get('master_sheet_id')
    #  update Sheet Id
    update_sheet_id = req.params.get('update_sheet_id')
    #  Master Target Column Name
    master_column_name = req.params.get('master_column_name')
    #  Update Target Dropdown Column Name
    update_column_name = req.params.get('update_column_name')

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

    logging.info(f"Column Info update_column_id:{updColumn_map[update_column_name]}")

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

    return func.HttpResponse("■ドロップダウンリストの更新が完了しました")


#############################################################################
###                                                                       ###
###  dropdownlist update timer                                            ###
###    parameters:master_sheet_id, update_sheet_id, master_column_name, update_column_name ###
###                                                                       ###
#############################################################################
@app.timer_trigger(schedule="0 */600 * * * *", arg_name="myTimer", run_on_startup=False,
              use_monitor=False) 
def dropdownlist_update_timer(myTimer: func.TimerRequest) -> None:

    logging.info('Python Timer trigger function processed a request.')
    logging.info(f"▼処理を開始します（dropdownlist_update_timer）")
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

#############################################################################
###                                                                       ###
###  webhook create                                                       ###
###    parameters:name, callback_url, sheet_id                            ###
###                                                                       ###
#############################################################################
@app.route(route="webhook_create", auth_level=func.AuthLevel.ANONYMOUS)
def webhook_create(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（webhook_create）")
    print("Start!!")

    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    try:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN_AZURE']
    except:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN']
    smart = smartsheet.Smartsheet(access_token)
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    # parameter
    name = req.params.get('name')
    callback_url_function = req.params.get('callback_url_function')
    function_list_sheet_id = req.params.get('function_list_sheet_id')
    sheet_id = int(req.params.get('sheet_id'))
    

    # callback_url
    callback_url = None
    funcSheet = smart.Sheets.get_sheet(function_list_sheet_id)
    for column in funcSheet.columns:
        funcColumn_map[column.title] = column.id
        #logging.info(f"cheack 2-1-2:{funcColumn_map[column.title]},{column.id}")

    # シート一覧を取得できている場合    
    if funcSheet:
        # シートの行数分ループ
        for row in funcSheet.rows:
            # functionの項目と引数の対象function名称が一致している項目を抽出
            function_name = ""
            for cell in row.cells:
                if cell.column_id == funcColumn_map['Function']:
                    function_name = cell.value
                    break

            if function_name == callback_url_function:
                # 一致した行のURL項目のURLを取得
                for cell in row.cells:
                    if cell.column_id == funcColumn_map['URL']:
                        callback_url = cell.value
                        logging.info(f"▼callback_url:{callback_url}")
                        break

    # master Sheet
    mstSheet = smart.Sheets.get_sheet(sheet_id)

    # Build column map for later reference - translates column names to column id
    for column in mstSheet.columns:
        mstColumn_map[column.title] = column.id
        
    column_id1 = None
    column_id2 = None
    column_id3 = None

    if req.params.get('column_name1'):
        column_id1 = mstColumn_map[req.params.get('column_name1')]
    if req.params.get('column_name2'):
        column_id2 = mstColumn_map[req.params.get('column_name2')]
    if req.params.get('column_name3'):
        column_id3 = mstColumn_map[req.params.get('column_name3')]

    try:
        logging.info("web_hook list get")
        # web_hook list get
        IndexResult = smart.Webhooks.list_webhooks(include_all=False)
        # target web_hook Delete
        for wh in IndexResult.data:
            if wh.name == name:
                smart.Webhooks.delete_webhook(wh.id)
                break

        if column_id1 and not column_id2 and not column_id3:
            webhook = smart.Webhooks.create_webhook(
                smartsheet.models.Webhook({
                    'name': name,
                    'callbackUrl': callback_url,
                    'scope': 'sheet',
                    'scopeObjectId': sheet_id,
                    'events': ['*.*'],
                    'version': 1,
                    'subscope': {
                        'columnIds': [column_id1],
                    }
                })
            )
        elif column_id1 and column_id2 and not column_id3:
            webhook = smart.Webhooks.create_webhook(
                smartsheet.models.Webhook({
                    'name': name,
                    'callbackUrl': callback_url,
                    'scope': 'sheet',
                    'scopeObjectId': sheet_id,
                    'events': ['*.*'],
                    'version': 1,
                    'subscope': {
                        'columnIds': [column_id1, column_id2],
                    }
                })
            )
        elif column_id1 and column_id2 and column_id3:
            webhook = smart.Webhooks.create_webhook(
                smartsheet.models.Webhook({
                    'name': name,
                    'callbackUrl': callback_url,
                    'scope': 'sheet',
                    'scopeObjectId': sheet_id,
                    'events': ['*.*'],
                    'version': 1,
                    'subscope': {
                        'columnIds': [column_id1, column_id2, column_id3],
                    }
                })
            )
        else:
            webhook = smart.Webhooks.create_webhook(
                smartsheet.models.Webhook({
                    'name': name,
                    'callbackUrl': callback_url,
                    'scope': 'sheet',
                    'scopeObjectId': sheet_id,
                    'events': ['*.*'],
                    'version': 1,
                })
            )

        if webhook.message == "SUCCESS":
            Webhook_upd = smart.Webhooks.update_webhook(
                webhook.data.id,       # webhook_id
                smart.models.Webhook({
                'enabled': True}))
        
        logging.info("web_hook list get")
        # web_hook list get
        IndexResult = smart.Webhooks.list_webhooks(
            page_size=100,
            page=1,
            include_all=False
        )

    except Exception as e:
        logging.error(f"Error creating webhook: {e}")
        return func.HttpResponse("■Webhookの作成が異常終了しました")

    logging.info(f"▲処理を終了します（webhook_create）")
    return func.HttpResponse("■Webhookの作成が完了しました")
