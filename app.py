from flask import Flask, request, abort
from linebot import  LineBotApi, WebhookHandler
from linebot.exceptions import  InvalidSignatureError
from linebot.models import *

#from datetime import datetime,timezone,timedelta
import datetime
import json,OriginalReply,MoneyReply



app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi('RY6oiNHZMTyvJFwpXuUVt2f5IpoM9pxcZOJzF+gwTwaLczarODcnNFt98B+auDkIYsZbiLDxnUTYzMVpf0Lg7F3zeVLVrLoU5kT5JFHBBHMGm+u6pLOHy0LhqV/0k2Q6cMK7P0KrHYu3KxCk0hUwZgdB04t89/1O/w1cDnyilFU=')
# Channel Secret
handler = WebhookHandler('e7ebf837ccbd2bacb20c9f90cea2ff0c')

previous_message = ""#記憶以前的訊息
#===============================================================================
# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

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
    
expenses_remaining=""
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    reply_arr=[]
    global previous_message
    print("獲取資料 : ",event.message.text,type(event.message.text))
    #Google試算表教學網頁 https://www.wongwonggoods.com/all-posts/python/python_web_crawler/python-pygsheets/
    auth_file = "linebotsheet.json"
    gc = pygsheets.authorize(service_file = auth_file)

    #時間設定
    dt1 = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc)
    dt2 = dt1.astimezone(datetime.timezone(datetime.timedelta(hours=8))) # 轉換時區 -> 東八區
    currentTime = dt2.strftime("%Y-%m-%d %H:%M:%S")
    
    if(event.message.text == '123'):   #獲取測試訊息
        reply_arr=OriginalReply.textReply(reply_arr,"獲取測試訊息")
    elif(event.message.text == 'q'):
        reply_arr=OriginalReply.quickReply(reply_arr)           
    elif(event.message.text == 'Google Sheet Test'):
        try:
            DataToGoogleSheet(currentTime,event.message.text,'Test')
            reply_arr=OriginalReply.textReply(reply_arr,"GoogleSheet上傳測試成功")
        except Exception as e:
            print("GoogleSheet上傳測試失敗 原因:",e)
            reply_arr=OriginalReply.textReply(reply_arr,"GoogleSheet上傳測試失敗")        
    
    elif(event.message.text == '當月剩餘費用'):
        expenses_remaining,RemainingCost=month_lessmoney(dt2,gc)
        reply_arr=OriginalReply.textReply(reply_arr,RemainingCost)
        reply_arr=OriginalReply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
        if expenses_remaining<=200:
            reply_arr=OriginalReply.textReply(reply_arr,"花太多錢啦!省錢一點")
        else:
            reply_arr=OriginalReply.textReply(reply_arr,"沒有超支 繼續保持!")
    elif(event.message.text == '當月信用卡費用'):
        LineBank=[]
        DaHo=[]
        datasheet,Month = MoneyGoogleSheet(dt2,gc)
        LineBank.append(datasheet.cell('E11').value)
        DaHo.append(datasheet.cell('E12').value)
        datasheet,Month = MoneyGoogleSheet(dt2,gc,1)
        LineBank.append(datasheet.cell('E9').value)
        DaHo.append(datasheet.cell('E10').value)
        reply_arr=OriginalReply.textReply(reply_arr,"LineBank信用卡 : "+str(int(LineBank[0]+LineBank[1]))+"元")
        reply_arr=OriginalReply.textReply(reply_arr,"永豐大戶信用卡 : "+str(int(DaHo[0]+DaHo[1]))+"元")

    elif(event.message.text == '本月記帳統計'):
        datasheet,Month = MoneyGoogleSheet(dt2,gc)
        MoneyType = datasheet.get_values_batch( ['K2:K11'])
        MoneyType = [item for sublist1 in MoneyType for sublist2 in sublist1 for item in sublist2]
        Money = datasheet.get_values_batch( ['L2:L11'])
        Money = [item for sublist1 in Money for sublist2 in sublist1 for item in sublist2]
        TotalMoney = datasheet.cell('D5')
        AllMoney = datasheet.cell('D3')
        reply_arr=MoneyReply.rankspend(reply_arr,AllMoney.value,TotalMoney.value,MoneyType,Money)
        
        
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
        
    elif(event.message.text == '汽機車格式'):
        datasheet = MotorGoogleSheet(dt2,gc)
        range_of_cells = datasheet.get_values_batch( ['B6:E6'])
        result_str = two_dimensional_list_intto_str(range_of_cells)
        reply_arr=OriginalReply.textReply(reply_arr,result_str)
    else:      
        try:
            #reply_arr=MoneyReply.MoneyquickReply(reply_arr,event.message.text)
            data_list = event.message.text.split(' ')
            try:
                outputtype = data_list[0]
                account = data_list[1]
                expendituretext = data_list[2]
                money = data_list[3]
            except:
                outputtype = "測試"
                account = "測試帳戶"
                expendituretext = "測試內容"
            money = "NT$ "+money
            print(data_list,money)
            if(outputtype == '收入' ):          
                try:
                    reply_arr=MoneyReply.expenditure(reply_arr,"新增收入成功",money,currentTime,outputtype,account,expendituretext)
                except:
                    reply_arr=MoneyReply.expenditure(reply_arr,"新增收入失敗",money,currentTime,outputtype,account,expendituretext)
            else:
                try:
                    reply_arr=MoneyReply.expenditure(reply_arr,"新增支出成功",money,currentTime,outputtype,account,expendituretext)
                except:
                    reply_arr=MoneyReply.expenditure(reply_arr,"新增支出失敗",money,currentTime,outputtype,account,expendituretext)
            
            DataToGoogleSheet(gc,dt2,data_list,'Money')

            datasheet,Month = MoneyGoogleSheet(dt2,gc)
            day=dt2.strftime("%d")
            TodayMoney = datasheet.cell('O'+str(int(day)))
            expenses_remaining,RemainingCost=month_lessmoney(dt2,gc)
                
            reply_arr=OriginalReply.textReply(reply_arr,"本日預算 : "+str("{:.2f}".format(expenses_remaining))+"元\n今天伙食費剩下 : "+str("{:.2f}".format((expenses_remaining)-int(TodayMoney.value)))+"元\n今天總花費"+str(TodayMoney.value)+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"記帳成功")
        except:      
            reply_arr=OriginalReply.textReply(reply_arr,"小企鵝壞掉了Q_Q")
            #textReply(reply_arr,event.message.text)
            if previous_message:
                reply_arr=OriginalReply.textReply(reply_arr,previous_message)
            else:
                reply_arr=OriginalReply.textReply(reply_arr,"目前還沒有前次訊息")
                previous_message = event.message.text
            #reply_arr=MoneyReply.MoneyquickReply(reply_arr,event.message.text)
       
    reply_arr=OriginalReply.DefaultQuickReply(reply_arr)    
    line_bot_api.reply_message(event.reply_token,reply_arr)     #LINE BOT回復訊息

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
