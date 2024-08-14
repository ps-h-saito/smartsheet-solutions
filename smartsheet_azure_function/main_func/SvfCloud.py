import logging
import json
import main_func as tr

clientId = "SVFQQVKUPULJHQVJRQLMRVMTRAVBGFTN"
secret = "oyi2EDxB2kxzUVibS1zWI6n7sYWPjBzka94YgWY72CJY3onL0tXr0gtkyFjflkzZ"
formFilePath = "form/Sample/見積書/見積書サンプル_注文書管理withsmartsheet.xml"

###
#  SVFCloudからPDFファイル取得
#  @param connect_str
#  @param input_file
#  @param output_file
#  @return 処理結果
###
def getSvfPdfData(connect_str,input_file,output_file):
    ### paramater ###
    # Authentication
    expDate = 300
    keyFilePath = "client.pkcs8"
    userId = "administrator"
    userName = "administrator"
    timeZone = "Asia/Tokyo"
    # PDF
    printerId = "PDF"
    #################

    try:
        logging.info(f"\n▼Accrss Token Get Start!---------------->\n") 
        # アクセストークン取得
        accessToken = tr.Authentication_getAccessToken(expDate,clientId,secret,keyFilePath,userId,userName,timeZone)
        logging.info(f"\n▲Accrss Token Get End!<------------------\n")
        logging.info(f"accessToken:{accessToken}")   
        res = False
        try:
            # PDF出力   
            logging.info(f"\n▼PDF Output Start!---------------->\n")         
            res = tr.PDF_pdfOutput(connect_str,accessToken,printerId,formFilePath,input_file,output_file)

            if res:
                json_printStatus = json.loads(res["printStatus"])
                if json_printStatus["state"] == 2:
                    logging.info(f"\n▲PDF Output Success!<--------------\n")
                else:
                    if res["download"]:
                        logging.info(f"\n▲PDF Output ファイルは作成されましたが、完了していません エラーコード：{json_printStatus['code']} <--------------\n")
                        logging.info(f"▲PDF Output File created but not completed! err_code:{json_printStatus['code']} <--------------\n")
                    else:
                        logging.info(f"\n▲PDF Output 帳票は作成されましたが、ファイル未作成です エラーコード：{json_printStatus['code']} <--------------\n")
                        logging.info(f"▲PDF Output The form has been created but the file has not been created  err_code:{json_printStatus['code']} <--------------\n")
            else:
                logging.info(f"▲PDF Output Error!<----------------\n")
        except Exception as e:
            logging.error(f"◆エラー発生 Exception:{e}")
        finally:
            logging.info(f"\n▼Accrss Token Revoke Start!---------------->\n") 
            # アクセストークン破棄
            if accessToken is not None:
                tr.Authentication_revokeAccessToken(accessToken)
                logging.info(f"\n▲Accrss Token Revoke End!<--------------\n")
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")