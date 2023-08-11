import pygsheets
# import xlrd
import pandas 
#for timezone
from dateutil.tz import tz
from datetime import datetime
from time import gmtime
# 時間套件 for shift month
import arrow


now_time = datetime.now()
#111

def get_week_alert():

   gc = pygsheets.authorize(service_file='E:\py_project\google_sheet\sheet_key.json')
   sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1YLDHZ8JLnzjJ024urc_YsOXpp1E-CJT-JeAPYRXEscQ/edit?usp=sharing')

   wks_list = sht.worksheets()
   # print(wks_list)
   sn=datetime.strftime(now_time,'%Y')+"出席"
   print(sn)
   wks2 = sht.worksheet_by_title(sn)


   # df = pandas.read_excel("E:\py_project\google_sheet\month_count.xlsx",sheet_name='工作表2')
   df = pandas.DataFrame(wks2.get_all_records())
   #set column names equal to values in row index position 0
   # df.columns = df.iloc[0]
   #remove first row from DataFrame
   # df = df[1:]

   # print("Columns")
   # print(df.columns)


   # for index, row in df.iterrows():
   #     print(row, end = "\n\n")

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
         print(row[activity_time_cn])
         time_list=row[activity_time_cn].split(" ")
         # print(time_list)
         activity_time=datetime.strptime(time_list[0]+" " +time_list[2] ,'%Y/%m/%d %H:%M')
         # print(activity_time)
         if now_time < activity_time:
            initiator=row[initiator_cn]
            palace=row[palace_cn]
            event=row[event_cn]
            w_d=week_list[activity_time.weekday()]
            show_date=datetime.strftime(activity_time,'%m/%d')+" "+w_d+" "+datetime.strftime(activity_time,'%H:%M')
            # print(activity_time)
            # print(week_list[activity_time.weekday()])
            # print(initiator)
            # print(palace)
            result_str+="-----------------------\n"
            result_str+="<"+str(initiator)+" 揪團中> " +event+" \n"
            result_str+="📅 "+show_date+' @'+palace+" \n"


   # print(result_str)         
   return result_str

def get_attend_dict():
   
   result_dict={}
   gc = pygsheets.authorize(service_file='E:\py_project\google_sheet\sheet_key.json')
   sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1YLDHZ8JLnzjJ024urc_YsOXpp1E-CJT-JeAPYRXEscQ/edit?usp=sharing')

   wks_list = sht.worksheets()
   # print(wks_list)
   sn_temp=datetime.strftime(now_time,'%Y')+"出席"
   ws_list=[sn_temp]

   # 若判斷上兩個月為跨年度，撈取上一年度分頁
   # tz_local = tzlocal()
   now_3_month = arrow.utcnow().shift(months=-3).datetime.replace(tzinfo=None)

   arw = now_3_month.strftime('%Y') 
   if datetime.strftime(now_time,'%Y') != arw:
      ws_list.append(arw+"出席")
   
   for sn in ws_list:
      wks2 = sht.worksheet_by_title(sn)


      # df = pandas.read_excel("E:\py_project\google_sheet\month_count.xlsx",sheet_name='工作表2')
      df = pandas.DataFrame(wks2.get_all_records())
      #set column names equal to values in row index position 0
      # df.columns = df.iloc[0]
      #remove first row from DataFrame
      # df = df[1:]

      # print("Columns")
      # print(df.columns)


      # for index, row in df.iterrows():
      #     print(row, end = "\n\n")

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
      # 月份
      month_cn = df.columns[5]

      user_list=df.columns.values[6:]


      for index, row in df.iterrows():
         # 讀本機檔案
         # if not pandas.isnull(row[activity_time_cn]):
         # 讀 google sheet
         if not row[activity_time_cn] == '':
            # print(row[activity_time_cn])
            time_list=row[activity_time_cn].split(" ")
            # print(time_list)
            activity_time=datetime.strptime(time_list[0]+"-"+gmtime().tm_zone,'%Y/%m/%d-%Z')
            
            act_time_str=datetime.strftime(activity_time,'%Y/%m')

            # print(activity_time)
            if now_3_month < activity_time:

               for user_cn in user_list:
                  user_name=row[user_cn]
                  # print(row[user_cn])
                  if user_name in result_dict:
                     result_dict[user_name].add(act_time_str)
                  else:
                     result_dict[user_name]={act_time_str}

               # initiator=row[initiator_cn]
               # palace=row[palace_cn]
               # event=row[event_cn]
               # w_d=week_list[activity_time.weekday()]
               # show_date=datetime.strftime(activity_time,'%m/%d')+" "+w_d+" "+datetime.strftime(activity_time,'%H:%M')
               # # print(activity_time)
               # # print(week_list[activity_time.weekday()])
               # # print(initiator)
               # # print(palace)
               # result_str+="-----------------------\n"
               # result_str+="<"+initiator+" 揪團中> " +event+" \n"
               # result_str+="📆 "+show_date+' @'+palace+" \n"


   # print(result_dict)         
   return result_dict


get_week_alert()
# attend_dict()