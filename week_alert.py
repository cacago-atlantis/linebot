# from google_sheet import get_week_alert
#google function
import json
import pygsheets
# import xlrd
import pandas 
from datetime import datetime
import requests    
import os


def push_msg():

# pyinstaller -F .\week_alert.py
    now_time = datetime.now()
    
    with open(os.getcwd()+"\config.json",encoding="utf-8") as json_file:
        config = json.load(json_file)
    
    group_type_text='測試機器人'
    # group_type_text='亞特'
    # group_type_text='團購'

    # test Channel Access Token
    token=config[group_type_text]['token']
    # group_id
    group_id=config[group_type_text]['group_id']

    


    # line_bot_api = LineBotApi(token)

    # 設定月初統計上傳圖片，路徑須為 Direct Link
    # static_url="https://i.imgur.com/VF0xxMK.png"
    # line_bot_api.push_message(group_id,ImageSendMessage(original_content_url=static_url, preview_image_url=static_url))

    msg_text=""

    msg_text+='hi hi 城民們:\n'    
    msg_text+='📖城民開的團歡迎參觀選購：\n'    
    # msg_text+=get_week_alert()

    gc = pygsheets.authorize(service_file=os.getcwd()+'\sheet_key.json')  
    # sheet_key_str = json.dumps(sheet_key_json)
    # gc = pygsheets.authorize(service_account_json=sheet_key_str)
    google_doc_path=config['google_doc_path']['城民清冊']
    # sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1YLDHZ8JLnzjJ024urc_YsOXpp1E-CJT-JeAPYRXEscQ/edit?usp=sharing')
    sht = gc.open_by_url(google_doc_path)

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