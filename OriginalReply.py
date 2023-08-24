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
