import requests
import time
import urllib
import urllib.parse
from base64 import b64encode
import json
from urllib.parse import urlparse
import logging
from azure.storage.blob import BlobServiceClient

### アクセストークン用定義 ###
API_DOMAIN = "https://api.svfcloud.com/" # APIを利用するためのドメイン
OAUTH2_TOKEN_URI = "oauth2/token" # アクセストークン取得のエンドポイント URI
OAUTH2_REVOKE_URI = "oauth2/revoke" # アクセストークン破棄のエンドポイント URI
RSA_ALGORITHM = "{\"alg\":\"RS256\"}" # 署名作成アルゴリズム（RSA using SHA-256 hash）
CLAIM_TEMPLATE = "{{\"iss\": \"{iss}\", \"sub\": \"{sub}\", \"exp\": \"{exp}\", \"userName\": \"{userName}\", \"timeZone\": \"{timeZone}\"}}" # 要求セットのテンプレート
GRANT_TYPE_JWT_BEARER = "urn:ietf:params:oauth:grant-type:jwt-bearer" # アクセストークン取得の認可タイプの値
PKCS8_HEADER = "-----BEGIN PRIVATE KEY-----" # 秘密鍵データヘッダー
PKCS8_FOOTER = "-----END PRIVATE KEY-----" # 秘密鍵データフッター
############################
### PDF出力用定義 ###
ACTIONS_URI = "v1/actions" # 帳票のを操作するエンドポイント URI
ARTIFACTS_URI = "v1/artifacts" # 帳票の成果物を操作するエンドポイント URI
CRLF = "\r\n" # 改行
############################


#######################################

###
#  アクセストークン取得
###
def Authentication_getAccessToken(expDate,clientId,secret,keyFilePath,userId,userName,timeZone):

    # JWT ベアラートークン（認証用トークン）を生成
    jwtToken = Authentication_generateJWTBearerToken(clientId, userId, keyFilePath, int(expDate), userName, timeZone)
    logging.info("jwtToken=\n " + jwtToken)

    # アクセストークンを取得
    accessToken = Authentication_getAccessTokenFromJWTBearerToken(jwtToken, clientId, secret)
    if accessToken is None:
        logging.info(f"◆エラー：アクセストークンを取得できませんでした。")
    else:
        logging.info("accessToken=\n " + accessToken)

    return accessToken

###
#  認証用トークン作成
#   
#  @param clientId
#  @param userId
#  @param keyFilePath
#  @param userName
#  @param timeZone
#  @return
###
def Authentication_generateJWTBearerToken(clientId, userId, keyFilePath, expDate, userName, timeZone):
    jwtToken = ""
    try:
        # 署名作成アルゴリズムを指定します。
        header = b64encode(bytes(RSA_ALGORITHM.encode('utf-8')))
        # 要求セットの作成
        payload = b64encode(bytes(Authentication_createPayload(clientId, userId, expDate, userName, timeZone).encode('utf-8')))
        # 署名対象のデータ
        jwtToken = jwtToken + header.decode() + "." + payload.decode()
        #/* 署名データを作成 */
        signedBytes = Authentication_signData(keyFilePath, str(jwtToken))
        # 署名データを付与して、JWTトークンの作成完了
        jwtToken = jwtToken + "." + signedBytes
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")
        return None

    return str(jwtToken)

###
#  認証用トークンからアクセストークン取得
#  @param jwtToken
#  @param clientId
#  @param secret
#  @return
###
def Authentication_getAccessTokenFromJWTBearerToken(jwtToken, clientId, secret):

    accessToken = None
    try:
        # パラメータ作成
        dict = {}
        dict["grant_type"] = GRANT_TYPE_JWT_BEARER
        dict["assertion"] = jwtToken
        payload = createPostParameter(dict)

        # httpヘッダーの作成
        authorization_data = (clientId + ":" + secret).encode('utf-8')
        authHeader = b64encode(bytes(authorization_data))
        headers = {"content-type": "application/x-www-form-urlencoded",
                   "Authorization": "Basic " + authHeader.decode()           
        }

        # リクエスト送信
        logging.info(f"-----TokenUriRequestInfo------\n payload=\n  [{payload}]\n headers=\n  [{headers}]\n-------------------------------")
        r = requests.post(API_DOMAIN + OAUTH2_TOKEN_URI, data=payload, headers=headers)
        logging.info(f"-----TokenUriResponseInfo------\n r.status_code:{r.status_code}\n r.url:{r.url}\n r.text:{r.text}\n r.encoding:{r.encoding}\n r.content:{r.content}\n-------------------------------")

        if r.status_code == 200:
            # 正常終了の場合リクエストの結果からアクセストークンを取り出します。
            try:
                logging.info(f"【{OAUTH2_TOKEN_URI} Response OK!】")
                accessToken = r.json().get('token')
            except:
               logging.error(f"◆Jsonデータからのアクセストークン取得でエラー発生")
        else:
            logging.info(f"【{OAUTH2_TOKEN_URI} Response NG!!】 r:[{r}]")
      
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")

    return accessToken

###
#  アクセストークン破棄
#  @param accessToken
###
def Authentication_revokeAccessToken(accessToken):

    try:
        # パラメータ作成
        dict = {}
        dict["token"] = accessToken
        payload = createPostParameter(dict)

        # httpヘッダーの作成
        headers = {"content-type": "application/x-www-form-urlencoded",
                   "Authorization": "Bearer " + accessToken
        }
        
        # リクエスト送信
        logging.info(f"------RevokeRequestInfo-------\n payload=\n  [{payload}]\n headers=\n  [{headers}]\n-------------------------------")
        r = requests.post(API_DOMAIN + OAUTH2_REVOKE_URI, data=payload, headers=headers)
        logging.info(f"------RevokeResponseInfo-------\n r.status_code:{r.status_code}\n r.url:{r.url}\n r.text:{r.text}\n r.encoding:{r.encoding}\n r.content:{r.content}\n-------------------------------")

        if r.status_code == 204:
            logging.info(f"【{OAUTH2_REVOKE_URI} Response OK!】")
        else:
            logging.info(f"【{OAUTH2_REVOKE_URI} Response NG!!】 r:[{r}]")

    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")

###
#  要求セットの作成
#  @param clientId
#  @param userId
#  @param expDate
#  @param userName
#  @param timeZone
#  @return 要求セット
###
def Authentication_createPayload(clientId, userId, expDate, userName, timeZone):

    now = int(time.time())
    return CLAIM_TEMPLATE.format(iss=clientId,sub=userId,exp=now+expDate,userName=userName,timeZone=timeZone)

###
#  署名データ作成
#  @param keyFilePath
#  @param signStr
#  @return 署名済みデータ
###
def Authentication_signData(keyFilePath, signStr):

    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.asymmetric import padding
    from cryptography.hazmat.primitives import hashes

    with open(keyFilePath, "rb") as key:
        p_key= serialization.load_pem_private_key(
            key.read(),
            password=None,
            backend=default_backend()
        )
    signature = p_key.sign(
        signStr.encode(),
        padding.PKCS1v15(),
        hashes.SHA256()
    )
    signature = b64encode(signature).decode()
    return signature

###
#  PDF出力
#  @param connect_str
#  @param accessToken
#  @param printerId
#  @param formFilePath
#  @param input_file
#  @param output_file
#  @return 出力結果(PDFダウンロードパス、成果物情報、印刷状況　帳票作成失敗はNone)
###
def PDF_pdfOutput(connect_str, accessToken,printerId,formFilePath,input_file,output_file):

    # 帳票出力
    location = SVF_print(connect_str, accessToken, printerId, formFilePath, input_file)
    if location is not None:
      logging.info("location=\n" + location)
    else:
      logging.info("print error.")
      return None

    # PDFダウンロード
    PDF_download(connect_str, accessToken, location, output_file)
    logging.info("output_file=\n" + output_file)

    # 成果物情報の取得
    artifactInfo = PDF_retrieveAtrifactInfo(accessToken, location)
    logging.info("artifactInfo=\n" + artifactInfo)

    # 印刷状況の取得
    printStatusLocation = API_DOMAIN + ACTIONS_URI + "/" + PDF_getActionId(location)
    printStatus = SVF_retrievePrintStatus(accessToken, printStatusLocation)
    logging.info("printStatus=\n" + printStatus)

    dict = {}
    dict["output_file"] = output_file
    dict["artifactInfo"] = artifactInfo
    dict["printStatus"] = printStatus

    return dict

###
#  PDFファイルダウンロード
#  @param connect_str
#  @param accessToken
#  @param location
#  @param output_file
###
def PDF_download(connect_str, accessToken, location, output_file):

    try :
        # httpヘッダーの作成
        headers = {"Accept": "application/octet-stream",
                   "Authorization": "Bearer " + accessToken
        }

        # リクエスト送信
        logging.info(f"------v1/artifactsRequestInfo_download-------\n headers=\n  [{headers}]\n------------------------------------")
        r = requests.get(location , headers=headers)
        logging.info(f"------v1/artifactsResponseInfo_download-------\n r.status_code:{r.status_code}\n r.url:{r.url}\n------------------------------------")

        if r.status_code == 200:
            # リクエストの結果からファイルを取り出して、出力先ファイルへ書き込みます。
            blob_service_client: BlobServiceClient = BlobServiceClient.from_connection_string(connect_str)
            blob_client = blob_service_client.get_blob_client(
                container='container-prd', blob=output_file)
            # ファイルが存在している場合は削除
            if blob_client.exists():
                blob_client.delete_blob()
            # ファイルの作成
            blob_client.upload_blob(r.content, blob_type="BlockBlob")
        elif r.status_code == 303:
            location = r.headers['Location']
            # 再度呼び出し先が指定された場合には、呼び出しを行います。
            PDF_download(connect_str, accessToken, location, output_file)
        else:
            logging.info(f"■bat status! [{r.status_code}][{r.text}]")
    
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")


###
#  成果物情報の取得
#  @param accessToken
#  @param location
#  @return 印刷状況
###
def PDF_retrieveAtrifactInfo(accessToken, location):
    artifactInfo = None
    
    try :
        # httpヘッダーの作成
        headers = {"Accept": "application/json",
                   "Authorization": "Bearer " + accessToken
        }

        # リクエスト送信
        logging.info(f"------v1/artifactsRequestInfo_AtrifactInfo-------\n headers=\n  [{headers}]\n-----------------------------------------")
        r = requests.get(location , headers=headers)
        logging.info(f"------v1/artifactsResponseInfo_AtrifactInfo-------\n r.status_code:{r.status_code}\n r.url:{r.url}\n r.text:{r.text}\n r.encoding:{r.encoding}\n r.content:{r.content}\n-----------------------------------------")

        if r.status_code == 200:
            artifactInfo = r.text
        else:
            logging.info(f"■bat status! [{r.status_code}][{r.text}]")
    
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")

    return artifactInfo

###
#  PDFダウンロード先のlocation から actionIdの取得
#  @param location
#  @return actionId
###
def PDF_getActionId(location):

    parsed_url = urlparse(location)
    queryList = str(parsed_url.query).split("&")
    dict = {}
    for query in queryList:
        split = query.split("=")
        dict[split[0]] = split[1]

    return dict.get("action")

###
#   * 帳票出力
#   * @param connect_str
#   * @param accessToken
#   * @param printerId
#   * @param formFilePath
#   * @param input_file
#   * @return
###
def SVF_print(connect_str, accessToken, printerId, formFilePath, input_file):

    location = None
    try:
        blob_service_client: BlobServiceClient = BlobServiceClient.from_connection_string(connect_str)
        blob_client_in = blob_service_client.get_blob_client(
                container='container-prd', blob=input_file)
        # ファイルデータの作成（data（辞書）の内容も含んで作成　※エラー回避のため）
        files = {
            "name": (None, "書類"),
            "source": (None, "CSV"),
            "printer": (None, printerId),
            "defaultForm": (None, formFilePath),
            # "password": (None, "psol"),
            "pdfPermPass": (None, "owner"),
            "pdfPermPrint": (None, "high"),
            "pdfPermModify": (None, "assembly"),
            "pdfPermCopy": (None, "true"),
            "redirect": (None, "false"),
            "data/書類": (input_file, blob_client_in.download_blob().readall(), "text/csv"),
        }

        # httpヘッダーの作成
        headers = {
            "Authorization": "Bearer " + accessToken,
        }

        # リクエスト送信
        logging.info(f"------v1/artifactsRequestInfo-------\n headers=\n  [{headers}]\n files=\n  [{files}]\n------------------------------------")
        r = requests.post(API_DOMAIN + ARTIFACTS_URI, headers=headers, files=files)
        logging.info(f"------v1/artifactsResponseInfo-------\n r.status_code:{r.status_code}\n r.url:{r.url}\n r.text:{r.text}\n r.encoding:{r.encoding}\n r.content:{r.content}\n------------------------------------")


        # 印刷の場合には、202、ファイルダウンロードの場合には、303のステータスコードが返ります。
        if r.status_code == 202 or r.status_code == 303:
            location = r.headers['Location']
            # locationには、
            # 印刷の場合->リクエストの結果から印刷状況を確認するURLが入ります。
            # ファイルダウンロードの場合->リクエストの結果からダウンロード先のURLが入ります。
        else:
            logging.info(f"■bat status! [{r.status_code}][{r.text}]")

    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")

    return location


###
#  印刷状況の取得
#  @param accessToken
#  @param location
#  @return 印刷状況
###
def SVF_retrievePrintStatus(accessToken, location):
    printerInfo = None

    try:
        # httpヘッダーの作成
        headers = {"Accept": "application/json",
                   "Authorization": "Bearer " + accessToken
        }

        # リクエスト送信
        logging.info(f"------v1/artifactsRequestInfo_PrintStatus-------\n headers=\n  [{headers}]\n----------------------------------------------")
        r = requests.get(location , headers=headers)
        logging.info(f"------v1/artifactsResponseInfo_PrintStatus-------\n r.status_code:{r.status_code}\n r.url:{r.url}\n r.text:{r.text}\n r.encoding:{r.encoding}\n r.content:{r.content}\n----------------------------------------------")

        if r.status_code == 200:
            # リクエストの結果から印刷状況を取り出します。
            printerInfo = r.text
        else:
            logging.info(f"■bat status! [{r.status_code}][{r.text}]")
    except Exception as e:
        logging.error(f"◆エラー発生 Exception:{e}")

    return printerInfo

###
#  Postパラメータ生成
#  @param paramMap
#  @return パラメータ
###
def createPostParameter(paramMap):
    sbPostParam = ""
    first = True
    for  k, v in paramMap.items():
        if first:
            first = False
        else:
            sbPostParam += "&"
        sbPostParam += urllib.parse.quote_plus(k) + "=" + urllib.parse.quote_plus(v)
    return sbPostParam
