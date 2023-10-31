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
def finding_Money_data(dt2,gc,DataArea):
    datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
    range_of_cells = datasheet.get_values_batch(DataArea)
    result_str = tool.two_dimensional_list_intto_str(range_of_cells)
    return result_str
#===============================================================================
previous_message = ""#記憶以前的訊息    
expenses_remaining=""
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    reply_arr=[]
    global previous_message
    print("獲取資料 : ",event.message.text,type(event.message.text))
    #Google試算表教學網頁 https://www.wongwonggoods.com/all-posts/python/python_web_crawler/python-pygsheets/
    gc = pygsheets.authorize(service_file = "linebotsheet.json")

    #時間設定
    dt1 = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc)
    dt2 = dt1.astimezone(datetime.timezone(datetime.timedelta(hours=8))) # 轉換時區 -> 東八區
    currentTime = dt2.strftime("%Y-%m-%d %H:%M:%S")
    try:
        if(previous_message == '記帳-支出'): 
            previous_message = ""
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
            #print(data_list,money)
            reply_arr=MoneyReply.expenditure(reply_arr,"新增支出成功",money,currentTime,outputtype,account,expendituretext)
                
            tool.DataToGoogleSheet(gc,dt2,data_list,'Money')
    
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            day=int(dt2.strftime("%d"))
            if day > 20:
              data_grid="A"+chr(int(day+44))
            else:
              data_grid=chr(int(ord('F'))+int(day))
                
            todayMoney = int(datasheet.cell(data_grid+'17').value)
            expenses_remaining,RemainingCost=tool.month_lessmoney(dt2,gc)
            
            reply_arr=Reply.textReply(reply_arr,"本日預算 : "+str("{:.2f}".format(expenses_remaining))+"元\n今天伙食費剩下 : "+str("{:.2f}".format((expenses_remaining)-int(todayMoney)))+"元\n今天總花費"+str(todayMoney)+"元")
            reply_arr=Reply.textReply(reply_arr,"記帳成功")
            
        elif(previous_message == '記帳-收入'): 
            previous_message = ""
            sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
            sheet = gc.open_by_url(sheet_url)
            try:
                datasheet = sheet.worksheet_by_title("總覽")
            except:
                print("沒有獲取到資料表")
                datasheet = sheet[0]
            data_list = event.message.text.split(' ')     
            values = [dt2.strftime("%m/%d")+str(data_list[0]),"收入"]+data_list[1:len(data_list)]
            try:
                A_column_values = datasheet.get_col(1, returnas='matrix', include_tailing_empty=False)
                datasheet.append_table(start='A' + str(len(A_column_values) + 1), end='F' + str(len(A_column_values) + 1), values=values)
            except Exception as e:
                print("寫入GoogleSheet error: ",e)
            reply_arr=Reply.textReply(reply_arr,"新增收入\n錢包 : "+str(data_list[1])+"\nLineBank : "+str(data_list[2])+"\n郵局 : "+str(data_list[3])+"\n永豐 : "+str(data_list[4]))
        elif(event.message.text == '記帳-支出'):
            print("進入記帳-支出模式")
            reply_arr=Reply.textReply(reply_arr,'進入記帳-支出模式')
            previous_message='記帳-支出'
            options = MoneyReply.lastest_four_data(dt2,gc)
            ReturnData = MoneyReply.lastest_four_data(dt2,gc,3)
            reply_arr.append(Reply.create_dropdown_menu(options,ReturnData))    
            
        elif(event.message.text == '記帳-收入'):
            reply_arr=Reply.textReply(reply_arr,'進入記帳-收入模式')
            reply_arr=Reply.textReply(reply_arr,'內容 錢包金額 LineBoank金額 郵局金額 永豐金額')           
            previous_message='記帳-收入'
        elif(event.message.text == '記帳-計畫'):
            reply_arr.append(Reply.creat_CarouselColumn(['建立新記帳','帳戶餘額','記帳類別','記帳帳號','記帳格式','當月剩餘費用','當月信用卡費用','本月記帳統計']))
        elif(event.message.text == '帳戶餘額'):
            print("帳戶餘額")    
            sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
            sheet = gc.open_by_url(sheet_url)
            datasheet = sheet.worksheet_by_title("總覽")
            balance = datasheet.get_values_batch(['C4:F4'])
            account = datasheet.get_values_batch(['C5:F5'])
            balance = tool.two_dimensional_list_intto_str(balance)
            account = tool.two_dimensional_list_intto_str(account)
            finial = []
            for i in range(len(account)):
                finial.append(account[i]+" : "+balance[i])
            print(finial)
            reply_arr=Reply.textReply(reply_arr,finial)
        elif(event.message.text == '當月剩餘費用'):
            expenses_remaining,RemainingCost=tool.month_lessmoney(dt2,gc)
            reply_arr=Reply.textReply(reply_arr,RemainingCost)
            reply_arr=Reply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
            if expenses_remaining<=200:
                reply_arr=Reply.textReply(reply_arr,"花太多錢啦!省錢一點")
            else:
                reply_arr=Reply.textReply(reply_arr,"沒有超支 繼續保持!")
        elif(event.message.text == '當月信用卡費用'):
            LineBank=[]
            DaHo=[]
            if(int(dt2.strftime("%d"))>=12):
                datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
                LineBank.append(datasheet.cell('E11').value)
                DaHo.append(datasheet.cell('E12').value)
                datasheet,Month = tool.MoneyGoogleSheet(dt2,gc,1)
                LineBank.append(datasheet.cell('E9').value)
                DaHo.append(datasheet.cell('E10').value)
            else:
                datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
                LineBank.append(datasheet.cell('E9').value)
                DaHo.append(datasheet.cell('E10').value)
                datasheet,Month = tool.MoneyGoogleSheet(dt2,gc,-1)
                LineBank.append(datasheet.cell('E11').value)
                DaHo.append(datasheet.cell('E12').value)
            reply_arr=Reply.textReply(reply_arr,"LineBank信用卡 : "+str(int(LineBank[0])+int(LineBank[1]))+"元")
            reply_arr=Reply.textReply(reply_arr,"永豐大戶信用卡 : "+str(int(DaHo[0])+int(DaHo[1]))+"元")
    
        elif(event.message.text == '本月記帳統計'):
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            MoneyType = datasheet.get_values_batch( ['K2:K11'])
            MoneyType = [item for sublist1 in MoneyType for sublist2 in sublist1 for item in sublist2]
            Money = datasheet.get_values_batch( ['L2:L11'])
            Money = [item for sublist1 in Money for sublist2 in sublist1 for item in sublist2]
            TotalMoney = datasheet.cell('D5')
            AllMoney = datasheet.cell('D3')
            reply_arr=MoneyReply.rankspend(reply_arr,AllMoney.value,TotalMoney.value,MoneyType,Money)
    
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            LineBank=(datasheet.cell('I2').value)
            DaHo=(datasheet.cell('I5').value)
            reply_arr=Reply.textReply(reply_arr,"LineBank信用卡 : "+str(LineBank)+"元")
            reply_arr=Reply.textReply(reply_arr,"永豐大戶信用卡 : "+str(DaHo)+"元")
                                 
        elif(event.message.text == '記帳類別'):
            result_str = finding_Money_data(dt2,gc,['K3:K11'])
            reply_arr=Reply.textReply(reply_arr,result_str)
        elif(event.message.text == '記帳帳號'):
            result_str = finding_Money_data(dt2,gc,['H2:H6'])
            reply_arr=Reply.textReply(reply_arr,result_str)
        elif(event.message.text == '記帳格式'):
            result_str = finding_Money_data(dt2,gc,['C14:F14'])
            reply_arr=Reply.textReply(reply_arr,result_str)
            
        elif(event.message.text == '汽機車格式'):
            datasheet = tool.MotorGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['B6:E6'])
            result_str = tool.two_dimensional_list_intto_str(range_of_cells)
            reply_arr=Reply.textReply(reply_arr,result_str)
        elif(event.message.text == '建立新記帳'): 
            Month = dt2.strftime("%m")
            sheet_url = "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=1747979925/"
            spreadsheet = gc.open_by_url(sheet_url)
            # 指定要複製的工作表名稱
            original_worksheet = spreadsheet.worksheet_by_title('記帳模板01')            
            # 複製工作表
            copied_worksheet = spreadsheet.add_worksheet(str(int(Month))+"月預算",src_worksheet=original_worksheet, index=2)
            reply_arr=Reply.textReply(reply_arr,'建立新記帳Finish')
        elif(event.message.text == 'test'):  
            print("Enter test")
            

        else:         
            reply_arr = Reply.ImageReply(reply_arr,tool.Image_searching(event.message.text))
            if previous_message:
                reply_arr=Reply.textReply(reply_arr,previous_message)
            else:
                reply_arr=Reply.textReply(reply_arr,"目前還沒有前次訊息")
                previous_message = event.message.text
    except Exception as e:      
        reply_arr=Reply.textReply(reply_arr,"小企鵝壞掉了Q_Q \n原因 : "+str(e))   
        
    if event.message.text != ('記帳-計畫') and event.message.text !=('記帳-支出'):  
        reply_arr.append(Reply.create_dropdown_menu(['記帳-支出','記帳-收入','記帳-計畫','test']))
    line_bot_api.reply_message(event.reply_token,reply_arr)     #LINE BOT回復訊息

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
