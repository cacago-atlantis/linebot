# from google_sheet import get_week_alert
#google function
import json
import pygsheets
# import xlrd
import pandas 
from datetime import datetime
import requests    

def push_msg():

# pyinstaller -F .\week_alert.py
    now_time = datetime.now()
    
    # # Channel Access Token
    # token='1m6m6T5ZJ9IFNJcodvQ3VnpjRLF4ccbBFh+bPEHL0Je8rP4l+NfFkngbSH30Tl6zftUFVOHjMABuanuXk/JfNUSqQag2gGdCGJBQ3sobG9MxDAWYlev8NggyEirQ3qya/GDA74wrfvoUOKtmc14UNAdB04t89/1O/w1cDnyilFU='
    # # 亞特
    # group_id='C14ab14ed596002ef031fe907f72f30cb'

    # test Channel Access Token
    token='ae19UEB3fm4w7mdYJ9k9xb6bNrjML1WXVaF89rqzqzswoCOxYY/Xcmpy6yreGcb3wpPnNc+4/YVT0idRouGHCEGDvNsXQrLWZMrRn++CoFoGbiWBDW0wQcMeRd0UEdox2SIGzHWxbIS/+IXWfClQMgdB04t89/1O/w1cDnyilFU='

    # 測試機器人
    group_id='C0899311c87a2d0f0e317f4ac82b3900c'
    # # 團購
    # group_id='C8f43d1d04b12c94a0d1e3ec315eedfad'

    # line_bot_api = LineBotApi(token)

    # 設定月初統計上傳圖片，路徑須為 Direct Link
    # static_url="https://i.imgur.com/VF0xxMK.png"
    # line_bot_api.push_message(group_id,ImageSendMessage(original_content_url=static_url, preview_image_url=static_url))

    msg_text=""

    msg_text+='hi hi 城民們:\n'    
    msg_text+='📖城民開的團歡迎參觀選購：\n'    
    # msg_text+=get_week_alert()

    # gc = pygsheets.authorize(service_file='C:\\Users\crar\Documents\project\py_project\google_sheet\sheet_key.json')
   
    sheet_key_json = {
  "type": "service_account",
  "project_id": "pristine-valve-387207",
  "private_key_id": "0b71001cdc3990396109e9ef926b991f51664fe5",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQDwPkdpVMxVcwSL\nVzUgWMGAlNjhOllv4/db1TneHUrznn4yvz+WnRET1O1v8/oA8pBBn9zISRrDORvA\naXza19aMLJfXpKjMT3SPgPqEdg7z/uwm7V+cdy9occouo6L3kgKsSxnrAxGyfbug\nv/EGgNEIm9N9Jpj7QboDaxFaBDlXusQdvE/l8BEMHlLOYT3i6tbyxFnKKvp2jz6K\nsUmsFmrFKt3lUr9khvCxlYKpzFC4cJ9okThmNC47sajbx+1notjGt2aEuTNdQj2G\nzs48rIofiIWToh4x1d0+S3ctUNyn3LuEl0zDDZypLhZVar5Qv45NQABjHZ6xm5de\nJoGR+l3ZAgMBAAECggEACjSbE/5mL7sTFlg+CYG7tpqcG/U+l2I3v87vBmB4LqEs\n+zrkmKFTeGQzMe5LUH79YcxItLlmSaTDaQkSJLnGg23hhWeZRBSb/vybp8TcHcFW\nhtDOnvbxLJ6o2BJbeejZ9De8gh8/QTXWIp4EvyH5v0PfWBGrrgC8xihmNEy5ouBX\nIXfMudFcUKLSQ+kbcml7xZNaAk1uc/AvNlxLdOHwe2cMmmF0AD0uBXBDII/p7/nm\nXXoEnLbuRMx1imQDFJkOTKZYGP7nsZhtLJxcsPltCq1mJC3GJZqRFIQDInKTR+71\nln53JnFJpfHM9vXiokmJnuhG2Q6lB802b6IBWDNOAQKBgQD85DJVETLJEM1Br/U4\nc5gnjQyWLvsz7gr/cwCbwT2/naurgM+9Owtx/8i71NZoLq4NtmKigoyIF5iNcrHY\no8DGqTUzbJCvxC2o2E/mQlOKN+g4uBmKNgZUK/j8pPUUaL2iGZxVe7xCfOvRfmKs\nFCmVNMmuPNbAA/grcC7c4K8bRQKBgQDzMkfwYigemkrqPzcdYxl5DA0it+1/yWCl\neq2R2ubjTodJzSZ3tSTZrBKcFuNJptUcya/dXLVaByhXGjsYYyKKO/poHt7jLyNo\nFJt54MYOPXiANOusrz1EAauFtXFTDOR6BRci9xf27uJehGDvpWXMLltPFZZ+dA2D\nwoZEHzYXhQKBgQDBGZX9MqautO+l6q+6LTnPaXpk6vbRTkCDkdKzG7kEqWY+DJuT\niJRStdcW5YvZ/VrWCaADKuAXwryvtRZrr44xo16GJ63LKGcc+B76WUbk0Y+2T4zg\n5iOq/fCfKW4h6WBzeE7RTywPMMf4LSM29iZSUf511uq7r8w9jumZqs7KaQKBgQCl\nyeA64l9hRWPOvtuOwBEMcQe/ZE2W8KxfAvuyU91UliMqT51qu+VsMp7ZI808V2wu\n3Ntz95B12C1K+8nPfT19qReyxWDC1U641FuNQYsjCArOs8T6CtikNNM+Kowfxsk2\n2aOFJZeDsiRFtM70b/eusudySVA30luoOAMaC4DvlQKBgQCkrs+2Ws+pCX4Eu+99\n9IDbqA6b98CGMIOWrfLgg7/MHOsKdkNle9LGg+OI7TQv3+BSkIPco2AOL4cRX/cu\nBqKoIVVcgmSiUTqF+7LIUUGAW7vLkazcB5IX7pGYQDigm0WRWOrKszysMnjFfQOD\nnu61lV7AbDNzJoj6K3aDmcezNg==\n-----END PRIVATE KEY-----\n",
  "client_email": "connect-google-sheet@pristine-valve-387207.iam.gserviceaccount.com",
  "client_id": "107311486915160569671",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/connect-google-sheet%40pristine-valve-387207.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}
    sheet_key_str = json.dumps(sheet_key_json)
    gc = pygsheets.authorize(service_account_json=sheet_key_str)
    sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1YLDHZ8JLnzjJ024urc_YsOXpp1E-CJT-JeAPYRXEscQ/edit?usp=sharing')

    wks_list = sht.worksheets()
    print(wks_list)
    # sn=datetime.strftime(now_time,'%Y')+"出席"
    # print(sn)
    sn='出席'
    wks2 = sht.worksheet_by_title(sn)

    df = pandas.DataFrame(wks2.get_all_records()).iloc[3:]

    result_str=""

    week_list = ["(一)","(二)","(三)","(四)","(五)","(六)","(日)"]
    # 時間
    activity_time_cn = df.columns[1]
    # 揪團者
    initiator_cn = df.columns[2]
    # 地點
    palace_cn = df.columns[3]
    # 活動名稱
    event_cn = df.columns[4]
    for index, row in df.iterrows():
        # 讀本機檔案
        # if not pandas.isnull(row[activity_time_cn]):
        # 讀 google sheet
        if not row[activity_time_cn] == '':
            time_list=row[activity_time_cn].split(" ")
        
            activity_time=datetime.strptime(time_list[0]+" " +time_list[2] ,'%Y/%m/%d %H:%M')

            if now_time < activity_time:
                initiator=row[initiator_cn]
                palace=row[palace_cn]
                event=row[event_cn]
                w_d=week_list[activity_time.weekday()]
                show_date=datetime.strftime(activity_time,'%m/%d')+" "+w_d+" "+datetime.strftime(activity_time,'%H:%M')

                result_str+="-----------------------\n"
                result_str+="<"+str(initiator)+" 揪團中> " +event+" \n"
                result_str+="📆 "+show_date+' @'+palace+" \n"

    msg_text+=result_str

    url = 'https://api.line.me/v2/bot/message/push'
    headers = {"Content-Type": "application/json; charset=utf-8",'Authorization': 'Bearer ' + token}
    myobj = { 'to': group_id, 
                "messages": [      
            {'type': "text",              
                "text": msg_text,}
                ]}

    x = requests.post(url,headers=headers, json = myobj)
    # push_url(msg_text)

    # try:
    #    line_bot_api.push_message(group_id, TextSendMessage(text=msg_text))

    # except LineBotApiError as e:
    #      # print(e)
    #      # error handle
    #      ...

push_msg()