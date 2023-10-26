from flask import Flask, request, abort
from linebot import  LineBotApi, WebhookHandler
from linebot.exceptions import  InvalidSignatureError
from linebot.models import *
from datetime import datetime, timedelta
import json,tool

#傳遞到GoogleSheet所使用的函示庫
import sys
import datetime
from oauth2client.service_account import ServiceAccountCredentials as SAC

app = Flask(__name__)

# 處理訊息
def textReply(reply_arr,Text):
    reply_arr.append(TextSendMessage(text=Text))
    return reply_arr
def ImageReply(reply_arr,url):
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
def DefaultQuickReply(options=None,textdata=None):
    '''
    if options == None:
        options = ["Option 1", "Option 2", "Option 3", "Option 4", "Option 5"]
    if textdata == None:
        textdata = options

    # 创建 QuickReplyButton 列表
    quick_reply_buttons = []
    
    for option in options:
        # 创建 MessageAction，这里将选项的文本作为消息
        action = MessageAction(label=option, text=option)
        
        # 创建 QuickReplyButton
        button = QuickReplyButton(action=action)
        
        # 将 QuickReplyButton 添加到列表
        quick_reply_buttons.append(button)
    
    # 创建 QuickReply 对象，包含 QuickReplyButton 列表
    quick_reply = QuickReply(items=quick_reply_buttons)
    
    # 创建 TextSendMessage 包含 QuickReply
    text_message = TextSendMessage(text="请选择一个选项：", quick_reply=quick_reply)

    return text_message
    '''
    text_message=TextSendMessage(
        text="快速回復選單",
        sticky=True,  # 将 sticky 参数设置为 True
        quick_reply=QuickReply(
            items=[ 
                QuickReplyButton(
                    action=MessageAction(label="記帳-支出",text="記帳-支出")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="記帳-收入",text="記帳-收入")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="記帳-計畫",text="記帳-計畫")
                    ),
                QuickReplyButton(
                    action=MessageAction(label="汽機車維修紀錄",text="汽機車維修紀錄")
                    )
                ]
            )
        )
    
    return text_message
    
#創造選單

def create_dropdown_menu(options=None,textdata=None):
    print("options : \n",options)
    print("textdata : \n",textdata)
    if options == None:
        options = ["三餐(食) 錢包 星巴克塑膠袋 3 ", "選項 2 ", "選項 3", "選項 4"]  # 这里可以根据你的需求设置选项
    if textdata == None:
        textdata = options
    actions = []           
    # 根据选项数量创建相应数量的按钮动作
    for i in range(len(options)):
        #print(len(options[i]),options[i])
        #print(len(textdata[i]),textdata[i])
        action = MessageTemplateAction(label=options[i],text=textdata[i])
        actions.append(action)
                
    # 创建 Buttons Template 消息
    buttons_template = ButtonsTemplate(text='請選擇選項',actions=actions)
                
    template_message = TemplateSendMessage(alt_text='下拉式選單',template=buttons_template)
    return template_message
    
def creat_CarouselColumn(options=None,textdata=None):
    # 选项列表，每个选项包含标题、描述和URL
    if options == None:
        options = ["Option 1", "Option 2", "Option 3", "Option 4", "Option 5", "Option 6", "Option 7"]
    if textdata == None:
        textdata = options
    # 创建 CarouselColumn 列表
    carousel_columns = []
    
    for i, option in enumerate(options):
        #print(option,textdata[i])
        # 创建 PostbackAction，这里将选项的文本作为 data 传递
        action = MessageAction(label=option, text=textdata[i])     
        # 创建 CarouselColumn
        column = CarouselColumn(
            thumbnail_image_url=tool.Image_searching(option),  # 缩略图的 URL
            title='Option',  # 列的标题
            text=option,  # 列的文本
            actions=[action]  # 列的行动
        )
        
        # 将 CarouselColumn 添加到列表
        carousel_columns.append(column)
    
    # 创建 CarouselTemplate
    carousel_template = CarouselTemplate(columns=carousel_columns)
    
    # 创建 TemplateSendMessage 包含 CarouselTemplate
    template_message = TemplateSendMessage(
        alt_text='Carousel Template',
        template=carousel_template,  # 显示 CarouselTemplate
        image_aspect_ratio='rectangle',  # 图像宽高比为矩形
        image_size='cover'  # 图像大小为覆盖整个区域
    )
    return template_message
