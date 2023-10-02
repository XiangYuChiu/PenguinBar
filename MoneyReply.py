from linebot.models import *
from datetime import datetime, timedelta
import json
import Reply,tool


def lastest_four_data(dt2,gc,x=4):
    worksheet,Month = tool.MoneyGoogleSheet(dt2,gc)    
    # 找到C列(3)到F列(6)非空白的数据的最后4笔新增数据
    non_empty_data=[]
    for i in range(x,7):
        non_empty_data.append([cell for cell in reversed(worksheet.get_col(i)) if cell.strip() != ""][:4])
    newest_four_data = []
    answer = ""
    for i in range(len(non_empty_data[0])):
        for j in range(len(non_empty_data)):
            answer += non_empty_data[j][i]+" "
        newest_four_data.append(answer) 
        answer = ""
    return newest_four_data


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

