import os

###
#  Smartsheetのアクセストークンを取得
#  @return アクセストークン
###
def get_smartsheet_Access_Token():

    access_token = None
    
    # 未加工のアクセストークンを環境変数から取得（新しいアクセストークンを生成で取得したコード）
    try:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN_AZURE']
    except:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN']


    # OAuth2でのアクセストークン取得




    return access_token
