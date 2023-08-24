from flask import Flask, request, abort
from linebot import  LineBotApi, WebhookHandler
from linebot.exceptions import  InvalidSignatureError
from linebot.models import *
from datetime import datetime, timedelta
import json

#傳遞到GoogleSheet所使用的函示庫
import sys
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials as SAC

app = Flask(__name__)
#===============================================================================
# 處理GoogleSheet
# 傳遞到GoogleSheet所使用的函示庫
import pygsheets
def MoneyGoogleSheet(dt2,gc,move=0):
    #print('MoneyGoogleSheet')
    sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
    sheet = gc.open_by_url(sheet_url)
    try:
        Month = dt2.strftime("%m")
        #print(str(int(Month))+"月預算")
        datasheet = sheet.worksheet_by_title(str(int(Month)+int(move))+"月預算")
    except:
        print("沒有獲取到資料表")
        datasheet = sheet[1]
    return datasheet,Month
def TestGoogleSheet(dt2,gc):  
    sheet_url = "https://docs.google.com/v/d/1nwrWBRMlfm-9oAnIeSXfuw5qKxqa4B4HCxCIUf36faM/edit#gid=0/"
    sheet = gc.open_by_url(sheet_url)
    datasheet = sheet[0]
    return datasheet

def MotorGoogleSheet(dt2,gc): 
    sheet_url = "https://docs.google.com/spreadsheets/d/1CKLgM4DdJ4njJEhMADszGo6skja0As520Lmgb6iwgcc/edit#gid=0"
    sheet = gc.open_by_url(sheet_url)
    datasheet = sheet.worksheet_by_title("MYN-7371") 
    return datasheet
    
def DataToGoogleSheet(gc,dt2,message,datatype): 
    #print("DataToGoogleSheet")
    if datatype == 'Test':
        datasheet=TestGoogleSheet(dt2,gc)
    elif datatype == 'Money':
        datasheet,Month = MoneyGoogleSheet(dt2,gc)
    elif datatype == 'Motor':
        datasheet=MotorGoogleSheet(dt2,gc)      
    values =  [dt2.strftime("%d")]+message
    print(values,type(values))
    try:
        datasheet.append_table(values)#這一行資料輸入完整 但是會失敗0727
    except Exception as e:
        print(e)
#===============================================================================
#===============================================================================
#工具類
def two_dimensional_list_intto_str(range_of_cells):
    # 使用循环将嵌套列表转换为字符串
    result_str = ''
    for sublist in range_of_cells:
        for inner_list in sublist:
            for item in inner_list:
                result_str += item + ' '
        
    # 去除字符串末尾多余的空格
    result_str = result_str.strip()
    return result_str
    
def create_dropdown_menu(options):
    actions = []
    if options == None:
        options = ['選項 1', '選項 2', '選項 3', '選項 4',]
    for option in options:
        action = MessageTemplateAction(label=option, text=option)
        actions.append(action)
    
    buttons_template = ButtonsTemplate(title='請選擇一個選項',  text='請選擇地區',actions=actions)
    template_message = TemplateSendMessage(alt_text='下拉式選單', template=buttons_template)

    return template_message
#===============================================================================
def month_lessmoney(dt2,gc):
    datasheet,Month = MoneyGoogleSheet(dt2,gc)
    RemainingCost = datasheet.cell('D2')
    Remaining=int(RemainingCost.value)
    RemainingCost = str(int(Month))+"月剩餘伙食費 : "+str(RemainingCost.value)+"元"        
                
    day=dt2.strftime("%d")
    current_date = datetime.datetime.now()# 获取当前日期
    first_day_of_month = current_date.replace(day=1)# 获取当前月份的第一天
                    
    # 获取下个月的第一天
    if first_day_of_month.month == 12:
        next_month = first_day_of_month.replace(year=first_day_of_month.year + 1, month=1)
    else:
        next_month = first_day_of_month.replace(month=first_day_of_month.month + 1)
                    
    # 计算当前月份的总天数
    total_days_in_month = (next_month - first_day_of_month).days
    expenses_remaining=int(Remaining)/(int(total_days_in_month)-int(day))
    return expenses_remaining,RemainingCost
#===============================================================================    
# 處理訊息
def textReply(reply_arr,Text):
    reply_arr.append(TextSendMessage(text=Text))
    return reply_arr
def PictureReply(reply_arr,url):
    reply_arr.append(ImageSendMessage(original_content_url=url, preview_image_url=url))
    return reply_arr
def ViedoReply(reply_arr,url):
    reply_arr.append(VideoSendMessage(original_content_url=url, preview_image_url=url))
    return reply_arr
def SoundReply(reply_arr,url):
    reply_arr.append(AudioSendMessage(original_content_url=url, duration=100000))
    return reply_arr
def LocalReply(reply_arr,Title,address,Latitude,Longitude):
    reply_arr.append(LocationSendMessage(title=Title, address=address, latitude=Latitude, longitude=Longitude))
    return reply_arr
def DefaultQuickReply(reply_arr):
    reply_arr.append(TextSendMessage(
        text="快速回復選單",
        sticky=True,  # 将 sticky 参数设置为 True
        quick_reply=QuickReply(
            items=[ 
                QuickReplyButton(
                    action=MessageAction(label="當月剩餘費用",text="當月剩餘費用")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="當月信用卡費用",text="當月信用卡費用")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="本月記帳統計",text="本月記帳統計")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="記帳格式",text="記帳格式")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="汽機車格式",text="汽機車格式")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="記帳類別",text="記帳類別")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="記帳帳號",text="記帳帳號")
                    ),
                QuickReplyButton(
                    action=PostbackAction(label="Postback",data="回傳資料")
                    ),
                QuickReplyButton(
                    action=DatetimePickerAction(label="時間選擇",data="時間選擇",mode='datetime')
                    ),
                QuickReplyButton(
                    action=CameraAction(label="拍照")
                    ),
                QuickReplyButton(
                    action=CameraRollAction(label="相簿")
                    ),
                QuickReplyButton(
                    action=LocationAction(label="傳送位置")
                    )
                ]
            )
        ))
    return reply_arr
#===============================================================================    
def Remaining_charges_for_the_month(reply_arr,dt2,gc):
    expenses_remaining,RemainingCost=month_lessmoney(dt2,gc)
    reply_arr=OriginalReply.textReply(reply_arr,RemainingCost)
    reply_arr=OriginalReply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
    if expenses_remaining<=200:
        reply_arr=OriginalReply.textReply(reply_arr,"花太多錢啦!省錢一點")
    else:
        reply_arr=OriginalReply.textReply(reply_arr,"沒有超支 繼續保持!")
    return reply_arr
''' 
elif(event.message.text == '當月信用卡費用'):
            LineBank=[]
            DaHo=[]
            if(int(dt2.strftime("%d"))>=12):
                datasheet,Month = MoneyGoogleSheet(dt2,gc)
                LineBank.append(datasheet.cell('E11').value)
                DaHo.append(datasheet.cell('E12').value)
                datasheet,Month = MoneyGoogleSheet(dt2,gc,1)
                LineBank.append(datasheet.cell('E9').value)
                DaHo.append(datasheet.cell('E10').value)
            else:
                datasheet,Month = MoneyGoogleSheet(dt2,gc)
                LineBank.append(datasheet.cell('E9').value)
                DaHo.append(datasheet.cell('E10').value)
                datasheet,Month = MoneyGoogleSheet(dt2,gc,-1)
                LineBank.append(datasheet.cell('E11').value)
                DaHo.append(datasheet.cell('E12').value)
            reply_arr=OriginalReply.textReply(reply_arr,"LineBank信用卡 : "+str(int(LineBank[0])+int(LineBank[1]))+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"永豐大戶信用卡 : "+str(int(DaHo[0])+int(DaHo[1]))+"元")
    
elif(event.message.text == '本月記帳統計'):
            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            MoneyType = datasheet.get_values_batch( ['K2:K11'])
            MoneyType = [item for sublist1 in MoneyType for sublist2 in sublist1 for item in sublist2]
            Money = datasheet.get_values_batch( ['L2:L11'])
            Money = [item for sublist1 in Money for sublist2 in sublist1 for item in sublist2]
            TotalMoney = datasheet.cell('D5')
            AllMoney = datasheet.cell('D3')
            reply_arr=MoneyReply.rankspend(reply_arr,AllMoney.value,TotalMoney.value,MoneyType,Money)
    
            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            LineBank=(datasheet.cell('I2').value)
            DaHo=(datasheet.cell('I5').value)
            reply_arr=OriginalReply.textReply(reply_arr,"LineBank信用卡 : "+str(LineBank)+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"永豐大戶信用卡 : "+str(DaHo)+"元")
            
            
            
            
elif(event.message.text == '記帳類別'):
            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['K3:K11'])
            result_str = two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
elif(event.message.text == '記帳帳號'):
            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['H2:H7'])
            result_str = two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
elif(event.message.text == '記帳格式'):
            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['C14:F14'])
            result_str = two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
'''
