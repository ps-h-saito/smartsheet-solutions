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
