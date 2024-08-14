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
AZURE_CONNECT_STR = ".net" # Azureストレージコネクト定義
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

