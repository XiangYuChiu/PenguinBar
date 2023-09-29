from flask import Flask, request, abort
from linebot import  LineBotApi, WebhookHandler
from linebot.exceptions import  InvalidSignatureError
from linebot.models import *

#from datetime import datetime,timezone,timedelta
import datetime, json,pygsheets
import OriginalReply,MoneyReply,tool



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


previous_message = ""#記憶以前的訊息    
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
    try:
        if(previous_message == '記帳-支出'): 
            #reply_arr=MoneyReply.MoneyquickReply(reply_arr,event.message.text)
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
            try:
                reply_arr=MoneyReply.expenditure(reply_arr,"新增支出成功",money,currentTime,outputtype,account,expendituretext)
            except:
                reply_arr=MoneyReply.expenditure(reply_arr,"新增支出失敗",money,currentTime,outputtype,account,expendituretext)
                
            tool.DataToGoogleSheet(gc,dt2,data_list,'Money')
    
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            day=int(dt2.strftime("%d"))
            if day > 20:
              data_grid="A"+chr(int(day+44))
            else:
              data_grid=chr(int(ord('F'))+int(day))
                
            todayMoney = int(datasheet.cell(data_grid+'17').value)
            expenses_remaining,RemainingCost=tool.month_lessmoney(dt2,gc)
            
            reply_arr=OriginalReply.textReply(reply_arr,"本日預算 : "+str("{:.2f}".format(expenses_remaining))+"元\n今天伙食費剩下 : "+str("{:.2f}".format((expenses_remaining)-int(todayMoney)))+"元\n今天總花費"+str(todayMoney)+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"記帳成功")
        elif(previous_message == '記帳-收入'): 
            previous_message = ""
            reply_arr.append(tool.create_dropdown_menu(previous_message))
        elif(event.message.text == '記帳-支出'):
            reply_arr=OriginalReply.textReply(reply_arr,'進入記帳-支出模式')
            previous_message='記帳-支出'
            newest_four_data = MoneyReply.lastest_four_data(dt2,gc)
            
        elif(event.message.text == '記帳-收入'):
            reply_arr=OriginalReply.textReply(reply_arr,'進入記帳-收入模式')
            previous_message='記帳-收入'
        elif(event.message.text == '當月剩餘費用'):
            expenses_remaining,RemainingCost=tool.month_lessmoney(dt2,gc)
            reply_arr=OriginalReply.textReply(reply_arr,RemainingCost)
            reply_arr=OriginalReply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
            if expenses_remaining<=200:
                reply_arr=OriginalReply.textReply(reply_arr,"花太多錢啦!省錢一點")
            else:
                reply_arr=OriginalReply.textReply(reply_arr,"沒有超支 繼續保持!")
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
            reply_arr=OriginalReply.textReply(reply_arr,"LineBank信用卡 : "+str(int(LineBank[0])+int(LineBank[1]))+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"永豐大戶信用卡 : "+str(int(DaHo[0])+int(DaHo[1]))+"元")
    
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
            reply_arr=OriginalReply.textReply(reply_arr,"LineBank信用卡 : "+str(LineBank)+"元")
            reply_arr=OriginalReply.textReply(reply_arr,"永豐大戶信用卡 : "+str(DaHo)+"元")
                                 
        elif(event.message.text == '記帳類別'):
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['K3:K11'])
            result_str = tool.two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
        elif(event.message.text == '記帳帳號'):
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['H2:H7'])
            result_str = tool.two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
        elif(event.message.text == '記帳格式'):
            datasheet,Month = tool.MoneyGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['C14:F14'])
            result_str = tool.two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
            
        elif(event.message.text == '汽機車格式'):
            datasheet = tool.MotorGoogleSheet(dt2,gc)
            range_of_cells = datasheet.get_values_batch( ['B6:E6'])
            result_str = tool.two_dimensional_list_intto_str(range_of_cells)
            reply_arr=OriginalReply.textReply(reply_arr,result_str)
            
        elif(event.message.text == 'test'):  
            print("Enter test")
            
            options = MoneyReply.lastest_four_data(dt2,gc)
            print(options)
            
            # 根据选项数量创建相应数量的按钮
            #options =["三餐(食) 錢包 星巴克塑膠袋 3 ", "購物(樂) 錢包 五金行 190 ", "三餐(食) LineBank信用卡 星巴克 335 ", "三餐(食) 錢包 宇峰塔帕尼 -75 "]
            options = ["三餐(食) 錢包 星巴克塑膠袋 3 ", "購物(樂) LineBank信用卡 五金行 190 ", "选项 4", "选项 4"]  # 这里可以根据你的需求设置选项
                
            actions = []
                
            # 根据选项数量创建相应数量的按钮动作
            for option in options:
                action = MessageTemplateAction(label=option,text=f'{option}')
                actions.append(action)
                
            # 创建 Buttons Template 消息
            buttons_template = ButtonsTemplate(title='请选择一个选项',text='请从下面的选项中选择一个',actions=actions)
                
            template_message = TemplateSendMessage(alt_text='下拉式选单',template=buttons_template)
            reply_arr.append(template_message)
            '''
            actions = []
            options = ['選項 1', '選項 2', '選項 3', '選項 4',]
            print(options)
            for option in options:
                action = MessageTemplateAction(label=option, text=option)
                actions.append(action)
            
            buttons_template = ButtonsTemplate(title='請選擇一個選項',  text='請選擇功能',actions=actions)
            template_message = TemplateSendMessage(alt_text='下拉式選單', template=buttons_template)
                
            reply_arr.append(template_message)
            '''
            #reply_arr=tool.create_default_dropdown_menu(reply_arr)
        else:         
            if previous_message:
                reply_arr=OriginalReply.textReply(reply_arr,previous_message)
            else:
                reply_arr=OriginalReply.textReply(reply_arr,"目前還沒有前次訊息")
                previous_message = event.message.text
    except Exception as e:      
        reply_arr=OriginalReply.textReply(reply_arr,"小企鵝壞掉了Q_Q \n原因 : "+str(e))   
        
    reply_arr=tool.create_default_dropdown_menu(reply_arr)
    #reply_arr=OriginalReply.DefaultQuickReply(reply_arr)    
    line_bot_api.reply_message(event.reply_token,reply_arr)     #LINE BOT回復訊息

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
