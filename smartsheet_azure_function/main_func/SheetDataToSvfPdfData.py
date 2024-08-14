import smartsheet
import logging
import smartsheet.token
import main_func as tr
from azure.storage.blob import BlobServiceClient
from datetime import datetime as dt

### 定義 ###
PDF_NAME_CREATE_ITEM_LIST = ["伝票ID","顧客社名","書類区分"] # PDFファイル名称命名規約用項目名
PDF_FILE_NAME = "{}_{}_{}.pdf" # 添付用PDFの名称
JUDGE_COLUMN_NAME = "状況" # 処理起動判断を行うカラム名称（パラメータのシートIDのシートに存在するカラムの名称を指定）
LOOP_START_COLUMN_NAME = "商品名" # １レコード内のループ開始名称
LOOP_END_COLUMN_NAME = "単価" # １レコード内のループ終了名称
AZURE_CONNECT_STR1 = "DefaultEndpointsProtocol=https;"
AZURE_CONNECT_STR2 = "AccountName=rgsmartsheetcustoma8a62;"
AZURE_CONNECT_STR3 = "AccountKey=T9SNqDKGwo4er1///"
AZURE_CONNECT_STR4 = "F04PqhQKOYtaN7UlFjEcSl4T9ownjb45GNMjXBww23BKERedK7nMBz5sIQG+AStVKwL6A==;"
AZURE_CONNECT_STR5 = "EndpointSuffix=core.windows.net"
AZURE_CONNECT_STR = AZURE_CONNECT_STR1 + AZURE_CONNECT_STR2 + AZURE_CONNECT_STR3 + AZURE_CONNECT_STR4 + AZURE_CONNECT_STR5 # Azureストレージコネクト定義
SITUATION_JUGMENT_NAME = "1.新規作成(PDF添付待ち)" # 実行判断条件
INPUT_FILE_NAME = "inputFile.csv" # Azureストレージ格納用CSVファイル名
OUTPUT_FILE_NAME = "outputFile.pdf" # Azureストレージ格納用PDFファイル名

# 出力CSV用固定情報マッピング定義
fixed_field_list = [["顧客社名","ClientName"],["伝票ID","OrderID"],["受発注年月日","OrderDate"]]
# 出力CSV用可変情報マッピング定義
column_list = [
    ["商品名1","ItemName"],["数量1","Num"],["単価1","UnitPrice"],
    ["商品名2","ItemName"],["数量2","Num"],["単価2","UnitPrice"],
    ["商品名3","ItemName"],["数量3","Num"],["単価3","UnitPrice"],
    ["商品名4","ItemName"],["数量4","Num"],["単価4","UnitPrice"],
    ["商品名5","ItemName"],["数量5","Num"],["単価5","UnitPrice"],
    ["商品名6","ItemName"],["数量6","Num"],["単価6","UnitPrice"],
    ["商品名7","ItemName"],["数量7","Num"],["単価7","UnitPrice"],
    ["商品名8","ItemName"],["数量8","Num"],["単価8","UnitPrice"],
    ["商品名9","ItemName"],["数量9","Num"],["単価9","UnitPrice"],
    ["商品名10","ItemName"],["数量10","Num"],["単価10","UnitPrice"]]

newColumn_map = {}
insColumn_map = {}

###
#  カラムマップ作製
#  @param row
#  @param column_name
#  @return カラム情報
###
def get_cell_by_new_column_name(row, column_name):
    column_id = newColumn_map[column_name]
    return row.get_column(column_id)
def get_cell_by_ins_column_name(row, column_name):
    column_id = insColumn_map[column_name]
    return row.get_column(column_id)

###
#  SVFCloudRestメイン処理
#  @param sheet_id
#  @param rowIdList
#  @return 処理結果
###
def svf_cloud_rest_main(sheet_id,rowIdList):

    logging.info(f"▼処理を開始します（sheet_id:{sheet_id} rowIdList:{rowIdList}）")

    # smartsheetオブジェクトの取得
    smart = smartsheet.Smartsheet(tr.get_smartsheet_Access_Token())
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)
    # シートIDからシート情報取得
    sheet = smart.Sheets.get_sheet(sheet_id)
    # シート情報からカラムマップ作製
    for column in sheet.columns:
        insColumn_map[column.title] = column.id
    # CSV作成用のtempシートのmodel作成
    sheet_spec = smartsheet.models.Sheet({
    'name': 'tmpnewsheet',
    'columns': [
            {
                'title': fixed_field_list[0][1],
                'type': 'TEXT_NUMBER'
            }, {
                'title': fixed_field_list[1][1],
                'type': 'TEXT_NUMBER'
            }, {
                'title': fixed_field_list[2][1],
                'type': 'TEXT_NUMBER'
            }, {
                'title': 'ItemName',
                'primary': True,
                'type': 'TEXT_NUMBER'
            }, {
                'title': 'Num',
                'type': 'TEXT_NUMBER'
            }, {
                'title': 'UnitPrice',
                'type': 'TEXT_NUMBER'
            }
    ]
    })
    # シートの行数ループ。１行を商品数分の行に変換してCSVに出力後、PDFファイルの作成
    for row in sheet.rows:
        if not row.id in rowIdList:
            continue

        # 判断を行うカラムの内容
        cell = get_cell_by_ins_column_name(row, JUDGE_COLUMN_NAME)
        if cell.value == SITUATION_JUGMENT_NAME:
            logging.info(f"◆処理対象 id:{row.id}") 
            # logging.info(f"◆処理対象 row:{row}") 
            # logging.info(f"◆処理対象 rowNumber:{str(row.rowNumber)}") 

            #####  事前処理  #####
            blob_service_client: BlobServiceClient = BlobServiceClient.from_connection_string(AZURE_CONNECT_STR)
            blob_client_in = blob_service_client.get_blob_client(
                container='container-prd', blob=INPUT_FILE_NAME)
            # 入力ファイルが存在している場合は削除
            if blob_client_in.exists():
                blob_client_in.delete_blob()
            blob_client_out = blob_service_client.get_blob_client(
                container='container-prd', blob=OUTPUT_FILE_NAME)
            # 出力ファイルが存在している場合は削除
            if blob_client_out.exists():
                blob_client_out.delete_blob()

            # tempシートの作成
            res = smart.Home.create_sheet(sheet_spec)
            # tempシートのシートID取得
            temp_sheet_id = res.data.id
            # tempシートIDからシート情報取得
            newSheet = smart.Sheets.get_sheet(temp_sheet_id)
            # シート情報からtempシート用のカラムマップ作製
            for column in newSheet.columns:
                newColumn_map[column.title] = column.id

            #####  処理_⓵  #####
            # tempシートへの追加用データ作成
            insRows = create_temp_insert_data(row)
            # tempシートへデータ追加
            if insRows :
                for insRow in insRows:
                    response = smart.Sheets.add_rows(
                        temp_sheet_id,
                        [insRow])
            # tempシートIDからtempシートのオブジェクト取得
            newSheet = smart.Sheets.get_sheet(temp_sheet_id)
            # snartsheetのtempシート情報をdataframeとして取得、
            df = tr.simple_sheet_to_dataframe(newSheet)
            # dataframe情報をCSVデータに変換
            output = df.to_csv(index=False)
            # CSVデータをAzureのストレージにアップロード
            blob_client_in.upload_blob(output)

            # tempシートの削除
            if temp_sheet_id is not None:
                logging.info(f"◆一時作成シート削除 temp_sheet_id:{temp_sheet_id}") 
                smart.Sheets.delete_sheet(temp_sheet_id)

            #####  処理_⓶⓷⓸  #####
            # SVFCloudでCSVをPDFに変換
            connect_str = AZURE_CONNECT_STR
            r = tr.getSvfPdfData(connect_str,INPUT_FILE_NAME,OUTPUT_FILE_NAME)

            #####  処理_⓹  #####
            # PDF名称作成
            pdf_name = PDF_FILE_NAME.format(
                get_cell_by_ins_column_name(row, PDF_NAME_CREATE_ITEM_LIST[0]).value,
                get_cell_by_ins_column_name(row, PDF_NAME_CREATE_ITEM_LIST[1]).value,
                get_cell_by_ins_column_name(row, PDF_NAME_CREATE_ITEM_LIST[2]).value)
            
            # 添付ファイルリストを取得
            response = smart.Attachments.list_row_attachments(
            sheet_id,       # sheet_id
            row.id,       # row_id
            include_all=True)
            attachmentFlg = False
            for attachment in response.data:
                logging.info(f"attachment:{attachment}")
                # 一致している添付ファイルが存在している場合、バージョンをアップとしてPDFファイルを登録
                if attachment.name == pdf_name:
                    response = smart.Attachments.attach_new_version(
                        sheet_id,       # sheet_id
                        attachment.id,       # attachment_id
                        (pdf_name,
                        blob_client_out.download_blob().readall(),
                        'application/pdf')
                    )
                    attachmentFlg = True
                    break
            
            if attachmentFlg == False:
                # PDFファイルを行の添付ファイルとして登録
                res = smart.Attachments.attach_file_to_row(
                    sheet_id, 
                    row.id, 
                    (pdf_name,
                    blob_client_out.download_blob().readall(),
                    'application/pdf')
                )

            # 添付ファイルを登録完了した場合、状況を「2.部長承認待ち」に変更
            # Build new cell value
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = cell.column_id
            new_cell.value = "2.部長承認待ち"
            new_cell.strict = False

            # Build the row to update
            new_row = smartsheet.models.Row()
            new_row.id = row.id
            new_row.cells.append(new_cell)
            # Update rows
            updated_row = smart.Sheets.update_rows(
                sheet_id,
                [new_row])

###
#  参照元シートの行データから挿入用のデータを作成
#  @param source_row
#  @return 挿入用行データ
###
def create_temp_insert_data(source_row):
    rows = []
    # 参照元シートの対象カラム
    for item in column_list:
        # 商品名の項目の場合
        if item[0][0:len(LOOP_START_COLUMN_NAME)] == LOOP_START_COLUMN_NAME:
            row = smartsheet.models.Row()
            row.to_last = True
            cell = get_cell_by_ins_column_name(source_row, item[0])
            # 参照元の商品名が存在しない場合は以降データ無しとして終了
            if cell.value is None or cell.value == "":
                break
            # 固定フィールド情報
            # 顧客社名
            cell = get_cell_by_ins_column_name(source_row, fixed_field_list[0][0])
            row.cells.append({
                'column_id': newColumn_map[fixed_field_list[0][1]],
                'value': cell.value or "",
                'strict': False
            })
            # 案件ID
            cell = get_cell_by_ins_column_name(source_row, fixed_field_list[1][0])
            row.cells.append({
                'column_id': newColumn_map[fixed_field_list[1][1]],
                'value': cell.value or "",
                'strict': False
            })
            # 受発注年月日
            cell = get_cell_by_ins_column_name(source_row, fixed_field_list[2][0])
            tstr = cell.value
            if cell.value != "":
                tdate = dt.strptime(cell.value, '%Y-%m-%dT%H:%M:%S')
                tstr = tdate.strftime('%Y/%m/%d').replace('/0', '/')
            row.cells.append({
                'column_id': newColumn_map[fixed_field_list[2][1]],
                'value': tstr or "",
                'strict': False
            })
        
        # 参照元の情報をtempシート用に追加
        cell = get_cell_by_ins_column_name(source_row, item[0])
        row.cells.append({
            'column_id': newColumn_map[item[1]],
            'value': cell.value or "",
            'strict': False
        })
        # 終了項目の場合に作成した行を追加
        if item[0][0:len(LOOP_END_COLUMN_NAME)] == LOOP_END_COLUMN_NAME:
            rows.append(row)

    return rows

