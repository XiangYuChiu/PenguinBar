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
        datasheet.append_table(start='B' + str(last_row_index + 1), end='F' + str(last_row_index + 1), values=values)
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
def Image_searching(Imagename):
    msg = Imagename
    img_search = {'tbm': 'isch', 'q': msg}
    query = urllib.parse.urlencode(img_search)
    base  = "https://www.google.com/search?"
    url   = str(base+query)
       
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36'}
            
    res  = urllib.request.Request(url, headers=headers)
    con  = urllib.request.urlopen(res)
    data = con.read()
    
    pattern = '"(https://encrypted-tbn0.gstatic.com[\S]*)"'
    img_list = []
    
    for match in re.finditer(pattern, str(data, "utf-8")):
        if len(match.group(1)) < 150:
            img_list.append(match.group(1))
    
    random_img_url = img_list[random.randint(0, len(img_list)+1)]
    message = ImageSendMessage(original_content_url = random_img_url,preview_image_url    = random_img_url)
    return message
