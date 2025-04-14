from FinMind.data import DataLoader
from datetime import timedelta
import datetime
from datetime import datetime as DT
import pandas as pd
import time
import calendar
from typing import List
#import talib

api = DataLoader()
with open("token_.txt") as token_file:
    token = token_file.read()

api.login_by_token(api_token=token)
## 先產生日期list，格式yyyy-mm-dd
start_date = "2025-04-01"
end_date = "2025-04-14"#datetime.today().strftime("%Y-%m-%d")
date_ls  = pd.date_range(start_date,end_date).strftime("%Y-%m-%d").tolist()

def split_data_by_time_ranges(df, time_column='time'):
    """
    將資料按照指定的時間範圍分割
    """
    # 確保時間欄位是datetime格式
    df['time_obj'] = pd.to_datetime(df[time_column], format='%H:%M:%S').dt.time
    
    # 定義兩個大的時間範圍
    range1_start = DT.strptime('00:00:00', '%H:%M:%S').time()
    range1_end = DT.strptime('05:00:00', '%H:%M:%S').time()
    
    range2_start = DT.strptime('08:45:00', '%H:%M:%S').time()
    range2_end = DT.strptime('13:45:00', '%H:%M:%S').time()
    
    range3_start = DT.strptime('15:00:00', '%H:%M:%S').time()
    range3_end = DT.strptime('00:00:00', '%H:%M:%S').time()
    
    # 創建空字典來存儲結果
    result_parts = {
        'part1': [],  # 00:00 - 05:00
        'part2': [],   # 08:45 - 13:45
        'part3': []   # 15:00 - 00:00
    }
    
    # 創建小時間範圍的空列表
    result_hourly = []
    time_tick = []
    
    # 處理第一個時間範圍 (00:00 - 05:00)
    current_time = range1_start
    for i in range(5):  # 5小時
        next_time = add_hour_to_time(current_time)
        time_tick.append(f"{current_time}~{next_time}")
        # 為跨小時的時間範圍篩選資料
        if current_time < next_time:
            mask = (df['time_obj'] >= current_time) & (df['time_obj'] < next_time)
        else:  # 處理跨夜情況
            mask = (df['time_obj'] >= current_time) | (df['time_obj'] < next_time)
        
        # 存儲這一小時的資料
        hourly_data = df[mask].copy()
        result_hourly.append(hourly_data)
        result_parts['part1'].append(hourly_data)
        
        # 更新下一個時間範圍的起始時間
        current_time = next_time
    
    # 處理第二個時間範圍 (8:45 - 13:45)
    current_time = range2_start
    for i in range(5):  # 5小時
        next_time = add_hour_to_time(current_time)
        time_tick.append(f"{current_time}~{next_time}")
        # 為跨小時的時間範圍篩選資料
        if current_time < next_time:
            mask = (df['time_obj'] >= current_time) & (df['time_obj'] < next_time)
        else:  # 處理跨夜情況
            mask = (df['time_obj'] >= current_time) | (df['time_obj'] < next_time)
        
        # 存儲這一小時的資料
        hourly_data = df[mask].copy()
        result_hourly.append(hourly_data)
        result_parts['part2'].append(hourly_data)
        
        # 更新下一個時間範圍的起始時間
        current_time = next_time
    
    # 處理第三個時間範圍 (15:00 - 00:00)
    current_time = range3_start
    for i in range(9):  # 9小時
        next_time = add_hour_to_time(current_time)
        time_tick.append(f"{current_time}~{next_time}")
        # 為跨小時的時間範圍篩選資料
        if current_time < next_time:
            mask = (df['time_obj'] >= current_time) & (df['time_obj'] < next_time)
        else:  # 處理跨夜情況
            mask = (df['time_obj'] >= current_time) | (df['time_obj'] < next_time)
        
        # 存儲這一小時的資料
        hourly_data = df[mask].copy()
        result_hourly.append(hourly_data)
        result_parts['part3'].append(hourly_data)
        
        # 更新下一個時間範圍的起始時間
        current_time = next_time
    
    return result_parts, result_hourly, time_tick

def add_hour_to_time(time_obj):
    """
    將時間增加一小時，處理跨夜情況
    """
    dt = DT.combine(DT.today(), time_obj) + timedelta(hours=1)
    return dt.time()

def clean_Constract_Date(df:pd.DataFrame):
        """
        資料清洗

        Args:
            df (pd.DataFrame): 從FinMind取得的原始資料

        Returns:
            df (pd.DataFrame): 只有月結算的資料
        """
        df["ok_date"] = [ym if len(ym)==6 else pd.NA for ym in df["contract_date"]]
        df.dropna(how="any",axis=0,inplace=True)
        df.drop(["ok_date"],axis=1,inplace=True)
        df.reset_index(inplace=True,drop=True)
        return df

def date_to_6num(date_dt:datetime.date):
        """將日期擷取為年4碼+當月2碼

        Args:
            date_dt (datetime.date): 交易日期

        Returns:
            (str): 轉換成當年當月
        """
        m = ""
        if len(str(date_dt.month)) == 1:
            m = "0"+str(date_dt.month)
        else:
            m = str(date_dt.month)
        return str(date_dt.year)+m

def date_to_6num_nextMonth(date_dt:datetime.date):
        """將日期擷取為年4碼+下個月2碼

        Args:
            date_dt (datetime.date): 交易日期

        Returns:
            (str): 下個月年月6碼
        """
        y = ""
        m = ""
        if len(str(date_dt.month)) == 1:
            y = str(date_dt.year)
            if str(date_dt.month) != "9":
                m = "0"+str(int(date_dt.month)+1)
            else:
                m = "10"
        else:
            if str(date_dt.month) != "12":
                y = str(date_dt.year)
                m = str(int(date_dt.month)+1)
            else:
                y = str(int(date_dt.year)+1)
                m = "01"
        return y+m

def third_wednesday(year: int, month: int) -> datetime.date:
        """計算指定年份與月份的第三個星期三 也就是交割日"""
        # 獲取該月的所有星期三
        wednesdays = [day for day in range(1, 32) 
                    if day <= calendar.monthrange(year, month)[1] and datetime.date(year, month, day).weekday() == 2
                    ]
        
        # 返回第三個星期三
        return datetime.date(year, month, wednesdays[2])

def find_needed_third_wed(df:pd.DataFrame)->List[datetime.date]:
        """
        找到所有資料中，年份與月份的第三個星期三
        Args:
            df (pd.DataFrame): _description_

        Returns:
            list[datetime.date]: _description_
        """
        dates = df["日期"].tolist()
        weds = []
        for date in dates:
            year, month = int(date.year), int(date.month)
            third_wed = third_wednesday(year, month)
            if third_wed not in weds:
                weds.append(third_wed)
        weds.sort()
        return weds

def get_1TF(df:pd.DataFrame):
        """
        擷取近一資料
        20250331更動：
        小時k的部分，
        結算日當日15:00前的是當月算近月，
        15:00後的下個月算近月

        Args:
            df (pd.DataFrame): 只有月結算的資料

        Returns:
            df_1tf (pd.DataFrame): 只有近月的資料
        """
        #cols = df.columns.tolist()
        df["日期"] = pd.to_datetime(df["日期"])
        df["time"] = pd.to_datetime(df["time"], format='%H:%M:%S').dt.time
        weds = find_needed_third_wed(df)
        tf_1_ym = []
        
        d = (df.loc[0,"日期"])
        for i in range(len(weds)):
            w_day = pd.to_datetime(weds[i])
            if d.year == w_day.year and d.month == w_day.month:
                break
        if d == w_day: #第三個周三(結算日)
            temp_time_line = DT.strptime('15:00:00', '%H:%M:%S').time()
            temp_df1 = df[df["time"]<temp_time_line]
            temp_df2 = df[df["time"]>=temp_time_line]
            tf_1_ym += [date_to_6num(d)]*len(temp_df1)
            tf_1_ym += [date_to_6num_nextMonth(d)]*len(temp_df2)
        else:
            for ind in range(len(df)):
                d = (df.loc[ind,"日期"])
                if d < w_day: #第三個周三以前
                    tf_1_ym.append(date_to_6num(d))
                elif d > w_day: #第三個周三以後(不含第三個周三)
                    tf_1_ym.append(date_to_6num_nextMonth(d))
        '''        
        for ind in range(len(df)):
            d = (df.loc[ind,"日期"])
            if d < w_day: #第三個周三以前
                tf_1_ym.append(date_to_6num(d))
            elif d == w_day: #第三個周三(結算日)
                        t = df.loc[ind,"time"]
                        t = DT.strptime(t,"%H:%M:%S").time()
                
                        if t.hour < 15: # 如果是15:00以前，date_to_6num
                            tf_1_ym.append(date_to_6num(d))
                        else: # 15:00以後(含15:00)，date_to_6num_nextMonth
                            tf_1_ym.append(date_to_6num_nextMonth(d))
            else: #第三個周三以後(包含第三個周三)
                tf_1_ym.append(date_to_6num_nextMonth(d))
        '''       
        df["近一年月"] = tf_1_ym
        df_1tf = df[df["近一年月"]==df["contract_date"]]
        #df_1tf = df_1tf[cols]
        df_1tf.reset_index(inplace=True,drop=True)
        return df_1tf
'''
def getMACD_OSC(df:pd.DataFrame)->pd.DataFrame:
        """
        計算MACD、OSC,並記錄在df內

        Args:
            df (pd.DataFrame): 日K資料

        Returns:
            df (pd.DataFrame): 加了OSC的日k資料
        """
        df["Close"] = df["Close"].astype("float")
        close_prices = df['Close'].values
        macd, macd_signal, macd_hist = talib.MACD(
            close_prices,
            fastperiod=12,   # 快速線週期
            slowperiod=26,   # 慢速線週期
            signalperiod=9   # 信號線週期
        )
        
        df['OSC'] = macd_hist
        
        return df
'''
## 使用 期貨交易明細 API
all_df_ls = []
for date in date_ls:
    df = api.taiwan_futures_tick(
    futures_id='TX',
    date=date
    )
    if df.empty:
        continue
    df["日期"] = pd.to_datetime(df["date"].str.split(" ").str[0])
    df["日期"] = pd.to_datetime(df["日期"],format="%Y-%m-%d").dt.date
    df["time"] = df["date"].str.split(" ").str[1]
    df["time"] = pd.to_datetime(df["time"],format="%H:%M:%S").dt.time

    df = clean_Constract_Date(df)
    df = get_1TF(df)

    result_parts, result_hourly, time_ticks = split_data_by_time_ranges(df)

    new_data = []
    ind = 0
    for temp_df in result_hourly:
        temp_data = []
        if temp_df.empty:
            ind+=1
            print("該時段無交易")
            continue
        temp_data.append(temp_df["日期"].iloc[0])
        temp_data.append(time_ticks[ind])
        temp_data.append(temp_df["price"].iloc[0])
        temp_data.append(temp_df["price"].max())
        temp_data.append(temp_df["price"].min())
        temp_data.append(temp_df["price"].iloc[-1])
        #temp_data.append(len(temp_df))
        #temp_data.append(temp_df["volume"].sum())
        #temp_data.append(round(temp_df["volume"].sum()/1000))
        #print(temp_data)
        new_data.append(temp_data)
        ind+=1
    #new_df = pd.DataFrame(new_data,columns=["時間","開盤價","收盤價","最高價","最低價","成交筆數","成交量(股)","成交量(張)"])
    new_df = pd.DataFrame(new_data,columns=["日期","時間區間","Open","High","Low","Close"])
    print(new_df)
    all_df_ls.append(new_df)
    print("="*50)
all_df = pd.concat(all_df_ls,axis=0)
print(all_df.info())
df = all_df.copy()

#df = getMACD_OSC(df)
#data_start_time = "2022-01-01"
#df = df[df["日期"]>=pd.to_datetime(start_date)]
df.to_excel(f"{start_date}~{end_date}_小時k.xlsx",index=False)
print(df.tail(20))
print(df.info())
    

    