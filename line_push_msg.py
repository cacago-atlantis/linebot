from linebot import LineBotApi
from linebot.models import TextSendMessage
from linebot.models import ImageSendMessage
from linebot.exceptions import LineBotApiError
import calendar
from datetime import datetime
# from count_dialog import count_func
from google_sheet import get_week_alert

# pyinstaller -F .\line_push_msg.py
#111
# # Channel Access Token
token='1m6m6T5ZJ9IFNJcodvQ3VnpjRLF4ccbBFh+bPEHL0Je8rP4l+NfFkngbSH30Tl6zftUFVOHjMABuanuXk/JfNUSqQag2gGdCGJBQ3sobG9MxDAWYlev8NggyEirQ3qya/GDA74wrfvoUOKtmc14UNAdB04t89/1O/w1cDnyilFU='
# # 亞特
group_id='C14ab14ed596002ef031fe907f72f30cb'

# # test Channel Access Token
# token='ae19UEB3fm4w7mdYJ9k9xb6bNrjML1WXVaF89rqzqzswoCOxYY/Xcmpy6yreGcb3wpPnNc+4/YVT0idRouGHCEGDvNsXQrLWZMrRn++CoFoGbiWBDW0wQcMeRd0UEdox2SIGzHWxbIS/+IXWfClQMgdB04t89/1O/w1cDnyilFU='

# # 測試機器人
group_id='C0899311c87a2d0f0e317f4ac82b3900c'
# 團購
# group_id='C8f43d1d04b12c94a0d1e3ec315eedfad'
# content_type="月初統計"
content_type="每月群規提醒"
# content_type="每週活動公告"

# 設定月初統計上傳圖片，路徑須為 Direct Link
static_url="https://i.imgur.com/lMh3C7o.png"


line_bot_api = LineBotApi(token)

msg_text=""



if content_type=="每月群規提醒":

    now_time = datetime.now()
    lastDay=calendar.monthrange(now_time.year,now_time.month)[1]

    line_bot_api.push_message(group_id,ImageSendMessage(original_content_url=static_url, preview_image_url=static_url))
    msg_text+='[城內公告-每月群規提醒]\n'
    msg_text+='1.每月每個人發話量需達50句 \n'    
    msg_text+='等 ' + str(now_time.month)+"/"+str(lastDay)+' 過後\n'
    msg_text+='就正式統計' + str(now_time.month)+'月單月發話量\n'
    msg_text+='連續兩個月發話不到50句的人\n'
    msg_text+='就要含淚跟妳說掰掰囉\n'
    msg_text+='請留在我們身邊好嗎>_<\n'
    msg_text+='\n'
    msg_text+='2.考量大家上班忙碌，沒辦法看訊息，\n'
    msg_text+='如果當月有參加任何一團，\n'
    msg_text+='可直接折抵當月句數50句。\n'
    msg_text+='\n'
    msg_text+='3.若有不可抗力因素\n'
    msg_text+='導致當月不能一起跟大家聊天，\n'
    msg_text+='可以私訊小幫手請假(最多請一個月)，\n'
    msg_text+='當月發話量不列入計算。\n'
    msg_text+='\n'
    msg_text+='以上，有任何困擾或尷尬的地方\n'
    msg_text+='都歡迎找小幫手們私訊傾訴哦🤞'


#     <群規 ver_20230407>
# 亞特蘭提斯群規 
# 群規寬鬆，希望大家都能當個「有禮貌的人」
# 1.每個月每人發話量達50句，若該月發話量不足，有出席活動者可相抵，不列入未達標。
# 2.請勿自行更改「用戶自動加入」及「網站·行動條碼邀請」這兩個設定，感謝大家
# 3.若有需暫時離開群組，請向小幫手說明原因才有機會再回來喔
# 4.報名各項活動，請體諒主辦人員的辛苦，不要一直進進出出、猶豫不決。如非不得已無法參與，請務必提前告知主辦人員喔～
# 5.如有想設公告之訊息或需要幫忙分享按讚之連結等，煩請先私訊或直接於群組間先詢問小幫手們。



elif content_type=="月初統計":
    # msg_text+=str(count_func("pipe"))
    line_bot_api.push_message(group_id,ImageSendMessage(original_content_url=static_url, preview_image_url=static_url))

    msg_text=""
    msg_text+='Hihi 大家好我是亞特機器人🤖️\n'

    #不踢人
    # msg_text+='又到了統計發話量的時候\n'
    # msg_text+='這次很讚的沒有人要離開我們\n'
    # msg_text+='附上統計數據給大家參考'

    # 12月入群的新朋友們12月是蜜月期，請不要擔心哦
 
    # 要踢人   
    
    msg_text+='很遺憾統計這兩個月的發話量後\n' 
    msg_text+='有城民要離開我們了QQ\n'
    msg_text+='康、郡、叮董\n'
    msg_text+='謝謝妳們這陣子給我們的回憶\n'    
    msg_text+='希望以後還有機會和妳們相聚'

    # 12月入群的新朋友們12月是蜜月期，請不要擔心哦

elif content_type=="每週活動公告":
    msg_text+='hi hi 城民們:\n'    
    msg_text+='📖城民開的團歡迎參觀選購：\n'    
    msg_text+=get_week_alert()
    # msg_text+='<Kenix 揪團中> 唱歌了啦 \n'        
    # msg_text+='--- 2/19(日) @台北松江錢櫃\n'    
    # msg_text+='<南瓜 揪團中> 喝酒吃飯回高雄\n'  
    # msg_text+='--- 2/27(一) @Intention\n'  
    # msg_text+='<彈塗魚 揪團中> 台中分部-抱石啦：）\n'       
    # msg_text+='--- 3/19(日) @攀吶攀岩館'      
    
print(msg_text)





try:
    line_bot_api.push_message(group_id, TextSendMessage(text=msg_text))

except LineBotApiError as e:
    print(e)
    # error handle
    ...