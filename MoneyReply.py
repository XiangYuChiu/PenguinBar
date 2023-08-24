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
def Remaining_charges_for_the_month(reply_arr,dt2,gc):
    expenses_remaining,RemainingCost=month_lessmoney(dt2,gc)
    reply_arr=OriginalReply.textReply(reply_arr,RemainingCost)
    reply_arr=OriginalReply.textReply(reply_arr,"平均每日伙食費剩下 : "+str("{:.2f}".format(expenses_remaining))+"元")
    if expenses_remaining<=200:
        reply_arr=OriginalReply.textReply(reply_arr,"花太多錢啦!省錢一點")
    else:
        reply_arr=OriginalReply.textReply(reply_arr,"沒有超支 繼續保持!")    
    return reply_arr
 
def rankspend(reply_arr,AllMoney,TotalMoney,MoneyType,Money):
    reply_arr.append(FlexSendMessage(
            alt_text="本月記帳統計",     
            contents={
  "type": "bubble",
  "hero": {
    "type": "image",
    "url": "https://scdn.line-apps.com/n/channel_devcenter/img/fx/01_2_restaurant.png",
    "size": "full",
    "aspectRatio": "20:13",
    "aspectMode": "cover",
    "action": {
      "type": "uri",
      "uri": "https://linecorp.com"
    }
  },
  "body": {
    "type": "box",
    "layout": "vertical",
    "spacing": "md",
    "action": {
      "type": "uri",
      "uri": "https://linecorp.com"
    },
    "contents": [
      {
        "type": "text",
        "text": "總支出",
        "size": "xl",
        "weight": "bold"
      },
      {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "box",
            "layout": "baseline",
            "contents": [
              {
                "type": "text",
                "text": TotalMoney+" $",
                "weight": "bold",
                "size": "xl",
                "color": "#EA0000"
              },
              {
                "type": "text",
                "text": "已花費百分比:"+str(("{:.2f}".format(round((int(TotalMoney)/int(AllMoney))*100, 2))))+"%",
                "color": "#EA0000",
                "align": "end",
                "size": "xxs"
              }
            ]
          }
        ]
      },
      {
        "type": "box",
        "layout": "vertical",
        "spacing": "md",
        "action": {
          "type": "uri",
          "uri": "https://linecorp.com"
        },
        "contents": [
          {
            "type": "text",
            "text": "消費排行",
            "size": "xl",
            "weight": "bold"
          },
          {
            "type": "box",
            "layout": "vertical",
            "spacing": "sm",
            "contents": [
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[0],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[0]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[0])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[1],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[1]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[1])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[2],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[2]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[2])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                   {
                    "type": "text",
                    "text": MoneyType[3],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[3]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[3])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[4],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[4]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[4])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[5],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[5]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[5])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[6],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[6]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[6])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[7],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[7]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[7])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": MoneyType[8],
                    "weight": "bold",
                    "margin": "sm",
                    "flex": 0
                  },
                  {
                    "type": "text",
                    "text": Money[8]+" $",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  },
                  {
                    "type": "text",
                    "text": str(("{:.2f}".format(round((int(Money[8])/int(TotalMoney))*100, 2))))+"%",
                    "size": "sm",
                    "align": "end",
                    "color": "#aaaaaa"
                  }
                ]
              }
            ]
          }
        ]
      }
    ]
  },
  "footer": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "button",
        "style": "primary",
        "color": "#905c44",
        "margin": "xxl",
        "action": {
          "type": "uri",
              "label": "查看試算表",
              "uri": "https://docs.google.com/spreadsheets/d/1jnKkUIegnTrr1nA-fCCp9i-sOoiB3_of1Ry5uwUFSvI/edit#gid=878287969"
            }
          }
        ]
      }
    }))
    return reply_arr

def MoneyquickReply(reply_arr,money):
    reply_arr.append(TextSendMessage(
        text="支出還是收入呢",
        sticky=True,  # 将 sticky 参数设置为 True
        quick_reply=QuickReply(
            items=[           
                QuickReplyButton(
                    action=MessageAction(label="支出",text=money)
                    ),
                QuickReplyButton(
                    action=MessageAction(label="收入",text="-"+money)
                    )
                ]
            )
        ))
    return reply_arr
def expenditure(reply_arr,state,money,formatted_time,outputtype,account,expendituretext):
    reply_arr.append(FlexSendMessage(
            alt_text=state,     
            contents={
                "type": "bubble",
                "body": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                    "type": "text",
                    "text": state,
                    "color": "#09FF00",
                    "weight": "bold",
                    "size": "md"
                    },
                    {
                    "type": "text",
                    "text": money,
                    "weight": "bold",
                    "size": "xl",
                    "color": "#FF0000"
                    },
                    {
                    "type": "box",
                    "layout": "vertical",
                    "margin": "lg",
                    "spacing": "sm",
                    "contents": [
                        {
                        "type": "box",
                        "layout": "baseline",
                        "spacing": "sm",
                        "contents": [
                            {
                            "type": "text",
                            "text": "時間",
                            "color": "#aaaaaa",
                            "size": "sm",
                            "flex": 1
                            },
                            {
                            "type": "text",
                            "text": formatted_time,
                            "wrap": True,
                            "color": "#666666",
                            "size": "sm",
                            "flex": 5
                            }
                        ]
                        },
                        {
                        "type": "box",
                        "layout": "baseline",
                        "spacing": "sm",
                        "contents": [
                            {
                            "type": "text",
                            "text": "類別",
                            "color": "#aaaaaa",
                            "size": "sm",
                            "flex": 1
                            },
                            {
                            "type": "text",
                            "text": outputtype,
                            "wrap": True,
                            "color": "#666666",
                            "size": "sm",
                            "flex": 5
                            }
                        ]
                        },
                        {
                        "type": "box",
                        "layout": "baseline",
                        "spacing": "sm",
                        "contents": [
                            {
                            "type": "text",
                            "text": "帳戶",
                            "color": "#aaaaaa",
                            "size": "sm",
                            "flex": 1
                            },
                            {
                            "type": "text",
                            "text": account,
                            "wrap": True,
                            "color": "#666666",
                            "size": "sm",
                            "flex": 5
                            }
                        ]
                        },
                        {
                        "type": "box",
                        "layout": "baseline",
                        "spacing": "sm",
                        "contents": [
                            {
                            "type": "text",
                            "text": "內容",
                            "color": "#aaaaaa",
                            "size": "sm",
                            "flex": 1
                            },
                            {
                            "type": "text",
                            "text": expendituretext,
                            "wrap": True,
                            "color": "#666666",
                            "size": "sm",
                            "flex": 5
                            }
                        ]
                        },
                    ]
                    }
                ]
                },
                "footer": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [],
                "flex": 0
                }
            }
    ))
    return reply_arr

