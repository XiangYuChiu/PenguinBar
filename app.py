from flask import Flask, request, abort
from linebot import  LineBotApi, WebhookHandler
from linebot.exceptions import  InvalidSignatureError
from linebot.models import *
import datetime, json,pygsheets
import Reply,MoneyReply,tool

app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi('RY6oiNHZMTyvJFwpXuUVt2f5IpoM9pxcZOJzF+gwTwaLczarODcnNFt98B+auDkIYsZbiLDxnUTYzMVpf0Lg7F3zeVLVrLoU5kT5JFHBBHMGm+u6pLOHy0LhqV/0k2Q6cMK7P0KrHYu3KxCk0hUwZgdB04t89/1O/w1cDnyilFU=')
# Channel Secret
handler = WebhookHandler('e7ebf837ccbd2bacb20c9f90cea2ff0c')
#===============================================================================
def finding_Money_data(timer,GoogleSheet,DataArea):
    datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
    range_of_cells = datasheet.get_values_batch(DataArea)
    result_str = tool.two_dimensional_list_intto_str(range_of_cells)
    return result_str
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
previous_message = ""#記憶以前的訊息    
expenses_remaining=""

def Accounting_Expenses(reply_arr):
    print("進入記帳-支出模式")
    reply_arr=Reply.textReply(reply_arr,'進入記帳-支出模式')
    previous_message='記帳-支出'
    options = MoneyReply.lastest_four_data(timer,GoogleSheet)
    ReturnData = MoneyReply.lastest_four_data(timer,GoogleSheet,3)
    reply_arr.append(Reply.create_dropdown_menu(options,ReturnData))  
    return reply_arr
    
def Accounting_Income(reply_arr):
    previous_message='記帳-收入'
    reply_arr=Reply.textReply(reply_arr,'進入記帳-收入模式')
    reply_arr=Reply.textReply(reply_arr,'內容 錢包金額 LineBoank金額 郵局金額 永豐金額')    
    return reply_arr
    
def Accounting_Plan(reply_arr):
    reply_arr.append(Reply.creat_CarouselColumn(['建立新記帳','帳戶餘額','記帳類別','記帳帳號','記帳格式','當月剩餘費用','當月信用卡費用','本月記帳統計']))
    return reply_arr
    
def Create_New_Accounting(reply_arr):
    Month = timer.strftime("%m")
    sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
    spreadsheet = GoogleSheet.open_by_url(sheet_url)
    # 指定要複製的工作表名稱
    original_worksheet = spreadsheet.worksheet_by_title('記帳模板01')            
    # 複製工作表
    copied_worksheet = spreadsheet.add_worksheet(str(int(Month))+"月預算",src_worksheet=original_worksheet, index=2)
    reply_arr=Reply.textReply(reply_arr,'建立新記帳Finish')
    return reply_arr
    
def Account_Balance(reply_arr):
    print("帳戶餘額")    
    sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
    sheet = GoogleSheet.open_by_url(sheet_url)
    datasheet = sheet.worksheet_by_title("總覽")
    balance = datasheet.get_values_batch(['C4:F4'])
    balance = [item for sublist1 in balance for sublist2 in sublist1 for item in sublist2]
    account = datasheet.get_values_batch(['C5:F5'])
    account = [item for sublist1 in account for sublist2 in sublist1 for item in sublist2]
    print(len(account),balance,account)
    finial = ""
    for i in range(len(account)):
        finial+=(account[i]+" : "+balance[i]+"$"+"\n")
    print(finial)
    reply_arr=Reply.textReply(reply_arr,finial)
    #reply_arr.append(Reply.creat_CarouselColumn(finial))
    return reply_arr

def Accounting_Category(reply_arr):
    result_str = finding_Money_data(timer,GoogleSheet,['K3:K11'])
    reply_arr=Reply.textReply(reply_arr,result_str)
    return reply_arr
    
def Accounting_Account(reply_arr):
    result_str = finding_Money_data(timer,GoogleSheet,['H2:H6'])
    reply_arr=Reply.textReply(reply_arr,result_str)
    return reply_arr
    
def Accounting_Format(reply_arr):
    result_str = finding_Money_data(timer,GoogleSheet,['C14:F14'])
    reply_arr=Reply.textReply(reply_arr,result_str)
    return reply_arr
    
def Remaining_Expenses_for_the_Month(reply_arr):
    expenses_remaining,RemaininGoogleSheetost=tool.month_lessmoney(timer,GoogleSheet)
    reply_arr=Reply.textReply(reply_arr,RemaininGoogleSheetost)
    reply_arr=Reply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
    if expenses_remaining<=200:
        reply_arr=Reply.textReply(reply_arr,"花太多錢啦!省錢一點")
    else:
        reply_arr=Reply.textReply(reply_arr,"沒有超支 繼續保持!")
    return reply_arr
        
def CreditCard_Charges_for_the_Month(reply_arr):
    LineBank=[]
    DaHo=[]
    if(int(timer.strftime("%d"))>=12):
        datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
        LineBank.append(datasheet.cell('E11').value)
        DaHo.append(datasheet.cell('E12').value)
        datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet,1)
        LineBank.append(datasheet.cell('E9').value)
        DaHo.append(datasheet.cell('E10').value)
    else:
        datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
        LineBank.append(datasheet.cell('E9').value)
        DaHo.append(datasheet.cell('E10').value)
        datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet,-1)
        LineBank.append(datasheet.cell('E11').value)
        DaHo.append(datasheet.cell('E12').value)
    reply_arr=Reply.textReply(reply_arr,"LineBank信用卡 : "+str(int(LineBank[0])+int(LineBank[1]))+"元")
    reply_arr=Reply.textReply(reply_arr,"永豐大戶信用卡 : "+str(int(DaHo[0])+int(DaHo[1]))+"元")
    return reply_arr
    
def Accounting_Statistics_for_this_Month(reply_arr):
    datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
    MoneyType = datasheet.get_values_batch( ['K2:K11'])
    MoneyType = [item for sublist1 in MoneyType for sublist2 in sublist1 for item in sublist2]
    Money = datasheet.get_values_batch( ['L2:L11'])
    Money = [item for sublist1 in Money for sublist2 in sublist1 for item in sublist2]
    TotalMoney = datasheet.cell('D5')
    AllMoney = datasheet.cell('D3')
    reply_arr=MoneyReply.rankspend(reply_arr,AllMoney.value,TotalMoney.value,MoneyType,Money)
    
    datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
    LineBank=(datasheet.cell('I2').value)
    DaHo=(datasheet.cell('I5').value)
    reply_arr=Reply.textReply(reply_arr,"LineBank信用卡 : "+str(LineBank)+"元")
    reply_arr=Reply.textReply(reply_arr,"永豐大戶信用卡 : "+str(DaHo)+"元")
    return reply_arr
    
def Automobile_and_Motorcycle_Format(reply_arr):
    datasheet = tool.MotorGoogleSheet(timer,GoogleSheet)
    range_of_cells = datasheet.get_values_batch( ['B6:E6'])
    result_str = tool.two_dimensional_list_intto_str(range_of_cells)
    reply_arr=Reply.textReply(reply_arr,result_str)
    return reply_arr
    
#===============================================================================
def Previous_Accounting_Expenses(reply_arr,message):
    print("進入支出紀錄")
    previous_message = ""
    data_list = message.split(' ')
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
    #print(data_list,money)
    reply_arr=MoneyReply.expenditure(reply_arr,"新增支出成功",money,currentTime,outputtype,account,expendituretext)
                
    tool.DataToGoogleSheet(GoogleSheet,timer,data_list,'Money')
    
    datasheet,Month = tool.MoneyGoogleSheet(timer,GoogleSheet)
    day=int(timer.strftime("%d"))
    if day > 20:
        data_grid="A"+chr(int(day+44))
    else:
        data_grid=chr(int(ord('F'))+int(day))
                
    todayMoney = int(datasheet.cell(data_grid+'17').value)
    expenses_remaining,RemaininGoogleSheetost=tool.month_lessmoney(timer,GoogleSheet)
            
    reply_arr=Reply.textReply(reply_arr,"本日預算 : "+str("{:.2f}".format(expenses_remaining))+"元\n今天伙食費剩下 : "+str("{:.2f}".format((expenses_remaining)-int(todayMoney)))+"元\n今天總花費"+str(todayMoney)+"元")
    reply_arr=Reply.textReply(reply_arr,"記帳成功")
    return reply_arr
    
def Previous_Accounting_Income(reply_arr,message):
    print("進入收入紀錄")
    previous_message = ""
    sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
    sheet = GoogleSheet.open_by_url(sheet_url)
    try:
        datasheet = sheet.worksheet_by_title("總覽")
    except:
        print("沒有獲取到資料表")
        datasheet = sheet[0]
    data_list = message.split(' ')     
    values = [timer.strftime("%m/%d")+str(data_list[0]),"收入"]+data_list[1:len(data_list)]
    try:
        A_column_values = datasheet.get_col(1, returnas='matrix', include_tailing_empty=False)
        datasheet.append_table(start='A' + str(len(A_column_values) + 1), end='F' + str(len(A_column_values) + 1), values=values)
    except Exception as e:
        print("寫入GoogleSheet error: ",e)
    reply_arr=Reply.textReply(reply_arr,"新增收入\n錢包 : "+str(data_list[1])+"\nLineBank : "+str(data_list[2])+"\n郵局 : "+str(data_list[3])+"\n永豐 : "+str(data_list[4]))
    return reply_arr
#===============================================================================

callback_dict={
    "記帳-支出":Accounting_Expenses,
    "記帳-收入":Accounting_Income,
    "記帳-計畫":Accounting_Plan,
    "建立新記帳":Create_New_Accounting,
    "帳戶餘額":Account_Balance,
    "記帳類別":Accounting_Category,
    "記帳帳號":Accounting_Account,
    "記帳格式":Accounting_Format,
    "當月剩餘費用":Remaining_Expenses_for_the_Month,
    "當月信用卡費用":CreditCard_Charges_for_the_Month,
    "本月記帳統計":Accounting_Statistics_for_this_Month,
    "汽機車格式":Automobile_and_Motorcycle_Format
}
previous_dict={
    "記帳-支出":Previous_Accounting_Expenses,
    "記帳-收入":Previous_Accounting_Income,
}
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    global previous_message,GoogleSheet,timer
    reply_arr = []
    message = event.message.text
    print("獲取資料 : ",message,type(message))
    #Google試算表教學網頁 https://www.wongwonggoods.com/all-posts/python/python_web_crawler/python-pygsheets/
    GoogleSheet = pygsheets.authorize(service_file = "linebotsheet.json")

    #時間設定
    dt1 = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc)
    timer = dt1.astimezone(datetime.timezone(datetime.timedelta(hours=8))) # 轉換時區 -> 東八區
    currentTime = timer.strftime("%Y-%m-%d %H:%M:%S")
    try:
        if message == 'test':
            print("Enter test")
        elif previous_message != "":
            selected_function = previous_dict[previous_message]
            reply_arr = selected_function(reply_arr,message)
            #previous_dict.get(message)
        else:
            selected_function = callback_dict[message]
            reply_arr = selected_function(reply_arr)
            #callback_dict.get(message)  
    except Exception as e:      
        reply_arr=Reply.textReply(reply_arr,"小企鵝壞掉了Q_Q \n原因 : "+str(e))   
        
    if message != ('記帳-計畫') and message !=('記帳-支出'):  
        reply_arr.append(Reply.create_dropdown_menu(['記帳-支出','記帳-收入','記帳-計畫','test']))
    line_bot_api.reply_message(event.reply_token,reply_arr)     #LINE BOT回復訊息

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
