import azure.functions as func
import datetime
import json
import logging
import smartsheet
import os

app = func.FunctionApp()

@app.route(route="test", auth_level=func.AuthLevel.ANONYMOUS)
def test(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

# ---------------------------------------------------
    # challenge check
    logging.info('challenge check')
    request_json = req.get_json()
    logging.info(f'get_req:{request_json}')

    if request_json and "challenge" in request_json:
        logging.info('challenge response')
        return json.dumps({
            "smartsheetHookResponse": request_json['challenge']
        })
    logging.info('not challenge response')
# ---------------------------------------------------

    name = req.params.get('name')
    if not name:
        try:
            req_body = req.get_json()
        except ValueError:
            pass
        else:
            name = req_body.get('name')

    if name:
        return func.HttpResponse(f"Hello, {name}. This HTTP triggered function executed successfully.")
    else:
        return func.HttpResponse(
             "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.",
             status_code=200
        )
    
@app.route(route="test_webhook_create", auth_level=func.AuthLevel.ANONYMOUS)
def test_webhook_create(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    logging.info(f"▼処理を開始します（test_webhook_create）")
    print("Start!!")

    logging.info("test 1")
    # Initialize client. Uses the API token in the environment variable "SMARTSHEET_ACCESS_TOKEN"
    try:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN_AZURE']
    except:
        access_token = os.environ['SMARTSHEET_ACCESS_TOKEN']
    smart = smartsheet.Smartsheet(access_token)
    # Make sure we don't miss any error
    smart.errors_as_exceptions(True)

    logging.info("test 2")
    # parameter
    name = req.params.get('name')
    callback_url = req.params.get('callback_url')
    sheet_id = int(req.params.get('sheet_id'))

    logging.info(f"test 3 {name},{callback_url},str({sheet_id})")

    logging.info("test 4")
    logging.info("web_hook list get")
    # web_hook list get
    IndexResult = smart.Webhooks.list_webhooks(
      page_size=100,
      page=1,
      include_all=False
    )

    logging.info("test 5")
    try:
        webhook = smart.Webhooks.create_webhook(
        smartsheet.models.Webhook({
            'name': name,
            'callbackUrl': callback_url,
            'scope': 'sheet',
            'scopeObjectId': sheet_id,
            'events': ['*.*'],
            'version': 1}))
        
        logging.info("test 6")
        logging.info(f"webhook:{webhook}")

        if webhook.message == "SUCCESS":
            Webhook_upd = smart.Webhooks.update_webhook(
                webhook.data.id,       # webhook_id
                smart.models.Webhook({
                'enabled': True}))
        
        logging.info("test 7")
        logging.info(f"Webhook_upd:{Webhook_upd}")
        
        # request_json = Webhook_upd.result.get_json()

        logging.info("test 8")
        logging.info("ftest 8.1 {request_json}")

    except Exception as e:
        logging.info("test 99 error")
        logging.info(e.__class__.__name__)
        logging.info(e.args) 
        logging.info(e)
        logging.info(f"{e.__class__.__name__}: {e}")
    

    logging.info("web_hook list get")
    # web_hook list get
    IndexResult = smart.Webhooks.list_webhooks(
      page_size=100,
      page=1,
      include_all=False
    )

    print("End!!")
    logging.info(f"▲処理を終了します（test_webhook_create）")
    return("")
