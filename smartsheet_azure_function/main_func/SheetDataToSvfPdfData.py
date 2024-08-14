import smartsheet
import logging
import smartsheet.token
import main_func as tr
from azure.storage.blob import BlobServiceClient
from datetime import datetime as dt

### 定義 ###
PDF_NAME_CREATE_ITEM_LIST = ["伝票ID","顧客社名","書類区分"] # PDFファイル名称命名規約用項目名

