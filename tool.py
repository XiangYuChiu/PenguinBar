from linebot.models import *
import datetime


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
        print(str(int(Month))+"月預算")
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
        b_column_values = datasheet.get_col(2, returnas='matrix', include_tailing_empty=False)
        last_row_index = len(b_column_values)
        datasheet.append_table(start='B' + str(last_row_index + 1), end='B' + str(last_row_index + 1), values=values)
        #datasheet.append_table(values)#這一行資料輸入完整 但是會失敗0727
    except Exception as e:
        print("寫入GoogleSheet error: ",e)
#===============================================================================
#工具類

#時間設定
def datetime_seting():
    dt1 = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc)
    dt2 = dt1.astimezone(datetime.timezone(datetime.timedelta(hours=8))) # 轉換時區 -> 東八區
    currentTime = dt2.strftime("%Y-%m-%d %H:%M:%S")
    return dt2

#二維矩陣轉換成字串
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
#創造選單
def create_dropdown_menu(options=None,textdata=None):
    # 使用HTML标签来自定义按钮文本的样式，实现居中对齐
    if options == None:
        options = ["三餐(食) 錢包 星巴克塑膠袋 3 ", "选项 2 ", "选项 3", "选项 4"]  # 这里可以根据你的需求设置选项
    if textdata == None:
        textdata = options
    actions = []           
    # 根据选项数量创建相应数量的按钮动作
    for i in range(len(options)):
        #print(i,options[i],textdata[i])
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
        # 创建 PostbackAction，这里将选项的文本作为 data 传递
        action = PostbackAction(label=option, data=textdata[i])
        
        # 创建 CarouselColumn
        column = CarouselColumn(
            thumbnail_image_url='https://example.com/thumb.jpg',  # 缩略图的 URL
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
    '''
    template_message= TemplateSendMessage(alt_text='Carousel template',template=CarouselTemplate(
            columns=[
                CarouselColumn(
                    thumbnail_image_url='https://example.com/item1.jpg',
                    title='this is menu1',
                    text='description1',
                    actions=[
                        PostbackTemplateAction(
                            label='postback1',
                            text='postback text1',
                            data='action=buy&itemid=1'
                        ),
                        MessageTemplateAction(
                            label='message1',
                            text='message text1'
                        ),
                        URITemplateAction(
                            label='uri1',
                            uri='http://example.com/1'
                        )
                    ]
                ),
                CarouselColumn(
                    thumbnail_image_url='https://example.com/item2.jpg',
                    title='this is menu2',
                    text='description2',
                    actions=[
                        PostbackTemplateAction(
                            label='postback2',
                            text='postback text2',
                            data='action=buy&itemid=2'
                        ),
                        MessageTemplateAction(
                            label='message2',
                            text='message text2'
                        ),
                        URITemplateAction(
                            label='uri2',
                            uri='http://example.com/2'
                        )
                    ]
                )
            ]
        )
    )
    '''
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
