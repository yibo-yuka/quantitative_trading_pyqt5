# region 套件載入區
from PyQt5 import QtCore,QtWidgets,QtGui
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, 
                             QVBoxLayout, QHBoxLayout, QLabel, QListWidget, 
                             QListWidgetItem, QTableWidget, QTableWidgetItem,
                             QPushButton, QHeaderView, QMessageBox)
from FinMind.data import DataLoader
import talib
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import re
import datetime
import time
import sys
import calendar
import os
# endregion

class mainwin(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("FITXN_1TF")
        self.setWindowTitle("台指期近一 回測分析")
        self.setWindowIcon(QtGui.QIcon('Frank_TXicon_24x24.ico')) #TODO
        self.resize(2000,1200)
        self.top_def_ck1 = True
        self.top_def_ck2 = False
        self.bot_def_ck1 = True
        self.bot_def_ck2 = False
        
        self.Hlayout = QtWidgets.QHBoxLayout(self)
        # 創建分割器
        #self.splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        #self.Hlayout.addWidget(self.splitter)
        
        self.right_wid = QtWidgets.QWidget()
        self.left_wid = QtWidgets.QWidget()
        self.left_wid.setGeometry(0,0,500,1200)
        self.right_part = QtWidgets.QHBoxLayout(self.right_wid)
        self.left_part = QtWidgets.QVBoxLayout(self.left_wid)
        self.strategy_ui()
        self.ReportTablet_ui()
        self.Hlayout.addWidget(self.left_wid)
        self.Hlayout.addWidget(self.right_wid)
        self.Hlayout.setStretchFactor(self.left_wid,1)
        self.Hlayout.setStretchFactor(self.right_wid,3)
        #self.splitter.addWidget(self.left_wid)
        #self.splitter.addWidget(self.right_wid)
        # 存儲已打開的標籤頁
        self.opened_tabs = ["損益歷史紀錄"]
        self.settings_saved = False
        
    def receive_checkbox_states(self, states):
        # 接收子視窗傳來的 checkbox 狀態
        self.checkbox_states = states
    
    def strategy_ui(self):
        
        self.strategyLb = QtWidgets.QLabel(self)
        self.strategyLb.setObjectName("StrategyLabel")
        self.strategyLb.setText("策略設定")
        self.strategyLb.setStyleSheet("""
                                      QLabel {
                                          font-family: 微軟正黑體;
                                          font-size: 50px;
                                          color: blue;
                                          font-weight: bold;
                                      }
                                      """)
        
        #self.strategyLb.move(50,20)
        self.left_part.addWidget(self.strategyLb)
        self.left_wid_1 = QtWidgets.QWidget()
        self.left_wid_2 = QtWidgets.QWidget()
        self.left_wid_3 = QtWidgets.QWidget()
        self.left_wid_4 = QtWidgets.QWidget()
        self.left_wid_5 = QtWidgets.QWidget()
        self.left_wid_6 = QtWidgets.QWidget()
        self.left_Vlayout1 = QtWidgets.QHBoxLayout(self.left_wid_1)
        self.left_Vlayout2 = QtWidgets.QHBoxLayout(self.left_wid_2)
        self.left_Vlayout3 = QtWidgets.QHBoxLayout(self.left_wid_3)
        self.left_Vlayout4 = QtWidgets.QHBoxLayout(self.left_wid_4)
        self.left_Vlayout5 = QtWidgets.QHBoxLayout(self.left_wid_5)
        self.left_Vlayout6 = QtWidgets.QHBoxLayout(self.left_wid_6)
        self.tacticCheckbtn1 = QtWidgets.QCheckBox("(不分正負)頂背離",self)
        #self.tacticCheckbtn1.move(50,100)
        self.tacticCheckbtn1.setChecked(True)
        self.left_Vlayout1.addWidget(self.tacticCheckbtn1)
        self.tacticCheckbtn3 = QtWidgets.QCheckBox("(不分正負)底背離",self)
        #self.tacticCheckbtn3.move(50,220)
        self.tacticCheckbtn3.setChecked(True)
        self.left_Vlayout3.addWidget(self.tacticCheckbtn3)
        self.tacticCheckbtn5 = QtWidgets.QCheckBox("頂背離",self)
        #self.tacticCheckbtn4.move(50,280)
        self.tacticCheckbtn5.setChecked(True)
        self.left_Vlayout5.addWidget(self.tacticCheckbtn5)
        self.tacticCheckbtn6 = QtWidgets.QCheckBox("底背離",self)
        #self.tacticCheckbtn4.move(50,280)
        self.tacticCheckbtn6.setChecked(True)
        self.left_Vlayout6.addWidget(self.tacticCheckbtn6)
        self.tacticCheckbtn2 = QtWidgets.QCheckBox("頂背離止損",self)
        #self.tacticCheckbtn2.move(50,160)
        self.tacticCheckbtn2.setChecked(True)
        self.left_Vlayout2.addWidget(self.tacticCheckbtn2)
        self.tacticCheckbtn4 = QtWidgets.QCheckBox("底背離止損",self)
        #self.tacticCheckbtn4.move(50,280)
        self.tacticCheckbtn4.setChecked(True)
        self.left_Vlayout4.addWidget(self.tacticCheckbtn4)
        
        self.adjustBtn1 = QtWidgets.QPushButton("調整定義",self)
        #self.adjustBtn1.move(250,90)
        #TODO 子視窗函數完成後再開
        #self.adjustTopDefine_ui = AdjustTopReverseDefination_ui()
        self.adjustBtn1.clicked.connect(self.open_top_define_setting)
        self.left_Vlayout1.addWidget(self.adjustBtn1)
        self.adjustBtn2 = QtWidgets.QPushButton("不可調整",self)
        self.adjustBtn2.setDisabled(True)
        #self.adjustBtn2.move(250,150)
        self.left_Vlayout2.addWidget(self.adjustBtn2)
        self.adjustBtn3 = QtWidgets.QPushButton("調整定義",self)
        #self.adjustBtn3.move(250,210)
        #TODO 子視窗函數完成後再開
        #self.adjustBotDefine_ui = AdjustBotReverseDefination_ui()
        self.adjustBtn3.clicked.connect(self.open_bot_define_setting)
        self.left_Vlayout3.addWidget(self.adjustBtn3)
        self.adjustBtn4 = QtWidgets.QPushButton("不可調整",self)
        self.adjustBtn4.setDisabled(True)
        #self.adjustBtn4.move(250,270)
        self.left_Vlayout4.addWidget(self.adjustBtn4)
        self.adjustBtn5 = QtWidgets.QPushButton("不可調整",self)
        self.adjustBtn5.setDisabled(True)
        #self.adjustBtn4.move(250,270)
        self.left_Vlayout5.addWidget(self.adjustBtn5)
        self.adjustBtn6 = QtWidgets.QPushButton("不可調整",self)
        self.adjustBtn6.setDisabled(True)
        #self.adjustBtn4.move(250,270)
        self.left_Vlayout6.addWidget(self.adjustBtn6)
        
        self.left_part.addWidget(self.left_wid_1)
        self.left_part.addWidget(self.left_wid_2)
        self.left_part.addWidget(self.left_wid_3)
        self.left_part.addWidget(self.left_wid_4)
        self.left_part.addWidget(self.left_wid_5)
        self.left_part.addWidget(self.left_wid_6)
        
        self.settingLb = QtWidgets.QLabel(self)
        self.settingLb.setObjectName("BackTestSettingLabel")
        self.settingLb.setText("回測報告設定")
        self.settingLb.setStyleSheet("""
                                      QLabel {
                                          font-family: 微軟正黑體;
                                          font-size: 50px;
                                          color: blue;
                                          font-weight: bold;
                                      }
                                      """)
        self.left_part.addWidget(self.settingLb)
        
        self.left_wid_5 = QtWidgets.QWidget()
        self.left_Vlayout5 = QtWidgets.QHBoxLayout(self.left_wid_5)
        
        self.date_Lb_st = QtWidgets.QLabel(self)
        self.date_Lb_st.setObjectName("DateLabel_start")
        self.date_Lb_st.setText("開始日期")
        self.left_Vlayout5.addWidget(self.date_Lb_st)
        #self.date_Lb.move(50,50)
        
        self.left_wid_7 = QtWidgets.QWidget()
        self.left_Vlayout7 = QtWidgets.QHBoxLayout(self.left_wid_7)
        
        self.min_15_kbar_cb = QtWidgets.QCheckBox("15分K",self)
        self.min_15_kbar_cb.setEnabled(False)
        self.hr_1_kbar_cb = QtWidgets.QCheckBox("小時K",self)
        self.hr_1_kbar_cb.setEnabled(True)
        self.hr_1_kbar_cb.setChecked(True)
        self.left_Vlayout7.addWidget(self.min_15_kbar_cb)
        self.left_Vlayout7.addWidget(self.hr_1_kbar_cb)
        self.left_part.addWidget(self.left_wid_7)
        
        self.date_val_1 = QtWidgets.QDateEdit(self)
        #self.date_val.setGeometry(110,50,150,30)
        self.date_val_1.setDisplayFormat("yyyy/MM/dd")
        self.date_val_1.setDate(QtCore.QDate().currentDate())
        self.left_Vlayout5.addWidget(self.date_val_1)
        
        self.date_Lb_end = QtWidgets.QLabel(self)
        self.date_Lb_end.setObjectName("DateLabel_end")
        self.date_Lb_end.setText("結束日期")
        self.left_Vlayout5.addWidget(self.date_Lb_end)
        #self.date_Lb.move(50,50)
        self.date_val_2 = QtWidgets.QDateEdit(self)
        #self.date_val.setGeometry(110,50,150,30)
        self.date_val_2.setDisplayFormat("yyyy/MM/dd")
        self.date_val_2.setDate(QtCore.QDate().currentDate())
        self.left_Vlayout5.addWidget(self.date_val_2)
        
        self.left_part.addWidget(self.left_wid_5)
        
        self.left_wid_6 = QtWidgets.QWidget()
        self.left_Vlayout6 = QtWidgets.QHBoxLayout(self.left_wid_6)
        self.principal_Lb = QtWidgets.QLabel(self)
        self.principal_Lb.setText("本金")
        self.left_Vlayout6.addWidget(self.principal_Lb)
        self.principal_val = QtWidgets.QLineEdit(self)
        #self.principal_val.setWidth(500)
        self.left_Vlayout6.addWidget(self.principal_val)
        self.left_part.addWidget(self.left_wid_6)
        
        self.backtestBtn = QtWidgets.QPushButton("開始回測",self)
        self.backtestBtn.setObjectName("BacktestBtn")
        self.backtestBtn.setStyleSheet("""
                                      QPushButton {
                                          font-family: 微軟正黑體;
                                          font-size: 30px;
                                          color: blue;
                                          font-weight: bold;
                                      }
                                      """)
        #self.backtestBtn.move(50,400)
        self.backtestBtn.clicked.connect(self.backTesting)
        self.left_part.addWidget(self.backtestBtn)
        
        self.ReportBtn = QtWidgets.QPushButton("匯出報表",self)
        self.ReportBtn.setStyleSheet("""
                                      QPushButton {
                                          font-family: 微軟正黑體;
                                          font-size: 30px;
                                          color: black;
                                          font-weight: bold;
                                      }
                                      """)
        #self.ReportBtn.move(50,520)
        #self.ReportBtn.clicked.connect(self.saveReportToExcel)
        self.ReportBtn.clicked.connect(self.open_selector)
        self.left_part.addWidget(self.ReportBtn)
    
    def ReportTablet_ui(self):
        '''
        self.ReadReportBtn = QtWidgets.QPushButton("匯入報表")
        self.ReadReportBtn.setStyleSheet("""
                                      QPushButton {
                                          font-family: 微軟正黑體;
                                          font-size: 30px;
                                          color: blue;
                                          font-weight: bold;
                                      }
                                      """)
        #self.ReadReportBtn.move(600,20)
        self.right_part.addWidget(self.ReadReportBtn)
        '''
        self.RecordTable = QtWidgets.QTableWidget(self)
        #self.RecordTable.setGeometry(QtCore.QRect(600,100,1200,1000))
        self.RecordTable.setObjectName("RecordTable")
        self.RecordTable.setColumnCount(6)
        self.RecordTable.setHorizontalHeaderLabels(["淨利($)","淨利(%)","平均獲利\n虧損比($)","平均獲利\n虧損比(%)","最大區間虧損($)","最大區間虧損(%)"])
        
        self.tabwid = QtWidgets.QTabWidget(self)
        self.tabwid.setTabsClosable(True)
        self.tabwid.tabCloseRequested.connect(self.close_tab_func)
        self.tab1 = QtWidgets.QWidget(self)
        self.tab1_layout = QtWidgets.QVBoxLayout(self.tab1)
        self.tab1_layout.addWidget(self.RecordTable)
        self.tabwid.addTab(self.tab1,"損益歷史紀錄")
        self.right_part.addWidget(self.tabwid)
        self.his_df = pd.DataFrame([["","","","","",""]],columns=["淨利($)","淨利(%)","平均獲利\n虧損比($)","平均獲利\n虧損比(%)","最大區間虧損($)","最大區間虧損(%)"])
        self.df_dict = {}
        self.df_dict["損益歷史紀錄"] = self.his_df
        self.text_list = QtWidgets.QListWidget(self)
        self.text_list.setMaximumWidth(300)  # 限制列表寬度
        self.text_list.addItem("損益歷史紀錄")
        self.right_part.addWidget(self.text_list)
    
    def close_tab_func(self,index):
        self.tabwid.removeTab(index)
    
    def get_TX_data(self,startDate:str,endDate:str,futuresId = 'TX')->pd.DataFrame:
        """
        從FinMind取得特定期貨的特定交易日期區間資料

        Args:
            startDate (str): 最早交易日期
            endDate (str): 最晚交易日期
            futuresId (str, optional): 期貨名稱, Defaults to 'TX'.

        Returns:
            df(pd.DataFrame): 從FinMind取得特定期貨的特定交易日期區間資料
        """
        api = DataLoader()
        df = api.taiwan_futures_daily(
            futures_id=futuresId,
            start_date=startDate,
            end_date=endDate,
        )
        df.columns = ['日期', 'futures_id', 'contract_date', 'Open', 'High', 'Low', 'Close',
            'spread', 'spread_per', 'Volume', 'settlement_price', 'open_interest',
            'trading_session']
        #print(df.columns)
        #df.to_excel(f"{futuresId}_{startDate}_{endDate}_info.xlsx")
        return df

    def clean_Constract_Date(self,df:pd.DataFrame):
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

    def date_to_6num(self,date_dt:datetime.date):
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

    def date_to_6num_nextMonth(self,date_dt:datetime.date):
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

    def date_to_6num_next2Month(self,date_dt:datetime.date):
        """將日期擷取為年4碼+下下個月2碼

        Args:
            date_dt (datetime.date): 交易日期

        Returns:
            (str): 下下個月年月6碼
        """
        y = ""
        m = ""
        if len(str(date_dt.month)) == 1:
            y = str(date_dt.year)
            if str(date_dt.month) != "8" and str(date_dt.month) != "9":
                m = "0"+str(int(date_dt.month)+1)
            elif str(date_dt.month) == "8":
                m = "10"
            elif str(date_dt.month) == "9":
                m = "11"
        else:
            if str(date_dt.month) != "11" and str(date_dt.month) != "12":
                y = str(date_dt.year)
                m = str(int(date_dt.month)+1)
            elif str(date_dt.month) == "11":
                y = str(int(date_dt.year)+1)
                m = "01"
            elif str(date_dt.month) == "12":
                y = str(int(date_dt.year)+1)
                m = "02"
        return y+m

    def third_wednesday(self,year: int, month: int) -> datetime.date:
        """計算指定年份與月份的第三個星期三 也就是交割日"""
        # 獲取該月的所有星期三
        wednesdays = [day for day in range(1, 32) 
                    if day <= calendar.monthrange(year, month)[1] and datetime.date(year, month, day).weekday() == 2
                    ]
        
        # 返回第三個星期三
        return datetime.date(year, month, wednesdays[2])

    def find_needed_third_wed(self,df:pd.DataFrame)->list[datetime.date]:
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
            year, month = date.year, date.month
            third_wed = self.third_wednesday(year, month)
            if third_wed not in weds:
                weds.append(third_wed)
        weds.sort()
        return weds

    def get_1TF(self,df:pd.DataFrame):
        """
        擷取近一資料
        20250322更動：
        因為資料中，同一日期的after market已經是當天15:00開盤的，
        所以說在第三個星期三當天以前的交易，近月是當月，
        在第三個星期三當天開始，近月都是下個月。

        Args:
            df (pd.DataFrame): 只有月結算的資料

        Returns:
            df_1tf (pd.DataFrame): 只有近一的資料
        """
        cols = df.columns.tolist()
        df["日期"] = pd.to_datetime(df["日期"])
        weds = self.find_needed_third_wed(df)
        tf_1_ym = []
        for ind in range(len(df)):
            d = (df.loc[ind,"日期"])
            for i in range(len(weds)):
                w_day = pd.to_datetime(weds[i])
                if d.year == w_day.year and d.month == w_day.month: #當月結算日
                    if d < w_day: #第三個周三以前
                        tf_1_ym.append(self.date_to_6num(d))
                    else: #第三個周三以後(包含第三個周三)
                        tf_1_ym.append(self.date_to_6num_nextMonth(d))
                    '''
                    else:# 第三個週三當天交易
                        if df.loc[ind,"trading_session"] == "position":#盤中
                            tf_1_ym.append(self.date_to_6num(d))
                        else: #after_market 盤後
                            tf_1_ym.append(self.date_to_6num_nextMonth(d))
                    '''
        df["近一年月"] = tf_1_ym              
        #df["YM"] = [date_to_6num_nextMonth(d) for d in df["日期"]]
        #df["ConDYM"] = [d[:6] for d in df["contract_date"]]
        df_1tf = df[df["近一年月"]==df["contract_date"]]
        #df_1tf = df_1tf[cols]
        df_1tf.reset_index(inplace=True,drop=True)
        return df_1tf

    def get_daliy_k_bar_range(self,df:pd.DataFrame,startdate,enddate)->pd.DataFrame:
        """
        整理成日k資料
        20250322更動：
        因為資料中，同一日期的after market已經是當天15:00開盤的，
        所以夜盤不需要再轉換成隔天日期，直接畫成k棒
        Args:
            df (pd.DataFrame): 近一資料

        Returns:
            df (pd.DataFrame): 日k資料
        """
        
        dates_ls = df["日期"].unique().tolist()
        dates_ls.sort()
        '''
        k_bar_days = []
        for i in range(len(df)):
            if df.loc[i,"trading_session"] == "position":
                if df.loc[i,"日期"] != pd.to_datetime(startdate):
                    k_bar_days.append(df.loc[i,"日期"])
                else:
                    k_bar_days.append(pd.NA)
            elif df.loc[i,"trading_session"] == "aftermarket":
                
                if df.loc[i,"日期"] != pd.to_datetime(enddate):
                    k_bar_days.append(dates_ls[(dates_ls.index(df.loc[i,"日期"])+1)])
                else:
                    k_bar_days.append(pd.NA)
                    
            else:
                k_bar_days.append(pd.NA)
        df["k_bar_date"] = k_bar_days
        df.dropna(how="any",inplace=True)
        df = df.sort_values(by = "k_bar_date")
        #df = df.sort_values(by = "日期")
        date_gp = df.groupby("k_bar_date")
        uni_date = df["k_bar_date"].unique().tolist()
        '''
        date_gp = df.groupby("日期")
        uni_date = df["日期"].unique().tolist()
        daliy_date = []
        for d in uni_date:
            daliy_df = date_gp.get_group(d)
            Open = 0
            High = 0
            Low = 0
            Close = 0
            Volumn = 0
            daliy_df.reset_index(inplace=True)
            for i in range(len(daliy_df)):
                if daliy_df.loc[i,"trading_session"] == "after_market":
                    Open = daliy_df.loc[i,"Open"]
                if daliy_df.loc[i,"trading_session"] == "position":
                    Close = daliy_df.loc[i,"Close"]
                
            if (not daliy_df.empty) and (Open == 0):
                Open = daliy_df.loc[0,"Open"]
            if (not daliy_df.empty) and (Close == 0):
                Close = daliy_df.loc[0,"Close"]
            if len(daliy_df) == 2:
                High = max(daliy_df.loc[0,"High"],daliy_df.loc[1,"High"])
                Low = min(daliy_df.loc[0,"Low"],daliy_df.loc[1,"Low"])
                Volumn = daliy_df.loc[0,"Volume"]+daliy_df.loc[1,"Volume"]
            elif len(daliy_df) == 1:
                High = daliy_df.loc[0,"High"]
                Low = daliy_df.loc[0,"Low"]
                Volumn = daliy_df.loc[0,"Volume"]
            daliy_row = [d,Open,High,Low,Close,Volumn]
            daliy_date.append(daliy_row)

        df = pd.DataFrame(daliy_date,columns = ["日期","Open","High","Low","Close","Volume"])
        df.drop_duplicates(keep="first",inplace=True)
        df.set_index('日期', inplace=True)
        return df

    def getMACD_OSC(self,df:pd.DataFrame)->pd.DataFrame:
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
        #df['MACD'] = macd
        #df['MACD_Signal'] = macd_signal
        #macd_hist.insert(0,pd.NA)
        #macd_hist = macd_hist[:-1]
        df['OSC'] = macd_hist
        #temp_ls = df['OSC'].tolist()
        #temp_ls.insert(0,pd.NA)
        #temp_ls = temp_ls[:-1]
        #df['OSC'] = temp_ls
        return df

    def find_pos(self,df:pd.DataFrame)->pd.DataFrame:
        global top_ck,top_stop_ck,bot_ck,bot_stop_ck,positive_top_ck,negative_bot_ck
        """
        依照OSC起伏找出波峰、波谷,並記錄波峰時最高價、波谷時最低價
        20250321細節補充：
        是在20240325偵測到20240322是OSC波峰，0326進場(做空)

        Args:
            df (pd.DataFrame): 有OSC的日k資料

        Returns:
            df (pd.DataFrame): 加了波峰波谷資料的日k資料
        """
        #df.reset_index(inplace=True)
        
        pos_ls = ["" for i in range(len(df))]
        high_ls = []
        low_ls = []
        for i in range(len(df)):
            if i > 1 and i != len(df)-1:
            #if i > 2:
                val_minus1 = df.loc[i-2,"OSC"]# 波峰波谷候選 前一個點
                val = df.loc[i-1,"OSC"]# 波峰波谷候選 
                val_add1 = df.loc[i,"OSC"]# 波峰波谷候選 後一個點
                #print(val_minus1,val,val_add1)
                if top_ck:
                    if val>val_minus1 and val>val_add1:
                        pos_ls[i] = "波峰"
                    
                if positive_top_ck:
                    all_positive_bool_ls = [v>0 for v in [val_minus1,val,val_add1]]
                    positive_check = all(all_positive_bool_ls)
                    if val>val_minus1 and val>val_add1 and positive_check:
                        pos_ls[i] = "波峰"
                        
                if bot_ck:
                    if val<val_minus1 and val<val_add1:
                        pos_ls[i] = "波谷"
                
                if negative_bot_ck:
                    all_negative_bool_ls = [v<0 for v in [val_minus1,val,val_add1]]
                    negative_check = all(all_negative_bool_ls)
                    if val<val_minus1 and val<val_add1 and negative_check:
                        pos_ls[i] = "波谷"
                '''
                if val>val_minus1 and val>val_add1:
                    pos_ls.append("波峰")
                    high_ls.append(df.loc[i-1,"High"])
                    low_ls.append("")
                elif val<val_minus1 and val<val_add1:
                    pos_ls.append("波谷")
                    high_ls.append("")
                    low_ls.append(df.loc[i-1,"Low"])
                else:
                    pos_ls.append("")
                    high_ls.append("")
                    low_ls.append("")
                '''
            '''
            else:
                pos_ls.append("")
                high_ls.append("")
                low_ls.append("")
            '''
        high_ls = [df.loc[i-1,"High"] if pos_ls[i] == "波峰" else "" for i in range(len(df))]
        low_ls = [df.loc[i-1,"Low"] if pos_ls[i] == "波谷" else "" for i in range(len(df))]
        pos_ls.append("") # 平移後index = len(df)-1
        high_ls.append("") # 平移後index = len(df)-1
        low_ls.append("") # 平移後index = len(df)-1
        #在找到波峰/波谷時，已經是下一天，所以波峰/波谷要往前平移一天
        #最高價與最低價已經是抓前一天的值，一起平移就好
        df["OSC波峰波谷"] = pos_ls[1:]
        df["波峰最高價"] = high_ls[1:]
        df["波谷最低價"] = low_ls[1:]
        return df
        
    def read_position(self,df:pd.DataFrame)->pd.DataFrame:
        global top_ck,top_stop_ck,bot_ck,bot_stop_ck,positive_top_ck,negative_bot_ck
        
        """
        判斷背離與止損

        Args:
            df (pd.DataFrame): 有OSC波峰波谷的日k資料

        Returns:
            df (pd.DataFrame): 判斷背離與止損的資料
        """
        #先設置背離欄位
        df["訊號"] = ["" for i in range(len(df))]
        #OSC波峰資料index df_t
        df_t = df[df["波峰最高價"] != ""]
        top_ind_ls = df_t.index.tolist()
        #OSC波谷資料index df_b
        df_b = df[df["波谷最低價"] != ""]
        bot_ind_ls = df_b.index.tolist()
        #讀取df_t每個row"OSC波峰波谷"，如果OSC跟最高價趨勢相反，紀錄頂背離訊號在df
        for i in range(len(top_ind_ls)-1):
            if top_ck:
                if self.top_def_ck1:
                    if df.loc[top_ind_ls[i],"OSC"]>df.loc[top_ind_ls[i+1],"OSC"] and df.loc[top_ind_ls[i],"波峰最高價"]<df.loc[top_ind_ls[i+1],"波峰最高價"]:
                        df.loc[top_ind_ls[i+1],"訊號"] += "(不分正負)頂背離＆"
                if self.top_def_ck2:
                    #這個保留起來，放UI上提供選擇
                    if df.loc[top_ind_ls[i],"OSC"]<df.loc[top_ind_ls[i+1],"OSC"] and df.loc[top_ind_ls[i],"波峰最高價"]>df.loc[top_ind_ls[i+1],"波峰最高價"]:
                        df.loc[top_ind_ls[i+1],"訊號"] += "(不分正負)頂背離＆"
            else:
                break
        #讀取df_b每個row"OSC波峰波谷"，如果OSC跟最低價趨勢相反，紀錄底背離訊號在df
        for i in range(len(bot_ind_ls)-1):
            if bot_ck:
                if self.bot_def_ck1:
                    if df.loc[bot_ind_ls[i],"OSC"]<df.loc[bot_ind_ls[i+1],"OSC"] and df.loc[bot_ind_ls[i],"波谷最低價"]>df.loc[bot_ind_ls[i+1],"波谷最低價"]:
                        df.loc[bot_ind_ls[i+1],"訊號"] += "(不分正負)底背離＆"
                if self.bot_def_ck2:
                    #這個保留起來，放UI上提供選擇
                    if df.loc[bot_ind_ls[i],"OSC"]>df.loc[bot_ind_ls[i+1],"OSC"] and df.loc[bot_ind_ls[i],"波谷最低價"]<df.loc[bot_ind_ls[i+1],"波谷最低價"]:
                        df.loc[bot_ind_ls[i+1],"訊號"] += "(不分正負)底背離＆"
            else:
                break
            
        if positive_top_ck:
            positive_top_ls = [False for i in range(len(df))]
            osc_ls = df["OSC"].tolist()
            position_ls = df["OSC波峰波谷"].tolist()
            for i in range(1,len(position_ls)-1):
                p = position_ls[i]
                all_positive = all([osc_ls[i-1]>0,osc_ls[i]>0,osc_ls[i+1]>0])
                if p == "波峰" and all_positive:
                    positive_top_ls[i] = True
            df["正向頂背離"] = positive_top_ls
            
            temp_df = df[df["正向頂背離"]]
            posi_top_ind_ls = temp_df.index.tolist()
            for i in range(len(posi_top_ind_ls)-1):
                if df.loc[posi_top_ind_ls[i],"OSC"]>df.loc[posi_top_ind_ls[i+1],"OSC"] and df.loc[posi_top_ind_ls[i],"波峰最高價"]<df.loc[posi_top_ind_ls[i+1],"波峰最高價"]:
                        df.loc[posi_top_ind_ls[i+1],"訊號"] += "頂背離＆"
                #TODO 之後要把另一個方向的趨勢背離加上
            df.drop("正向頂背離",axis=1,inplace=True)
            
        if negative_bot_ck:
            negative_bot_ls = [False for i in range(len(df))]
            osc_ls = df["OSC"].tolist()
            position_ls = df["OSC波峰波谷"].tolist()
            for i in range(1,len(position_ls)-1):
                p = position_ls[i]
                all_negative = all([osc_ls[i-1]<0,osc_ls[i]<0,osc_ls[i+1]<0])
                if p == "波谷" and all_negative:
                    negative_bot_ls[i] = True
            df["負向底背離"] = negative_bot_ls
            
            temp_df = df[df["負向底背離"]]
            nega_bot_ind_ls = temp_df.index.tolist()
            for i in range(len(nega_bot_ind_ls)-1):
                if df.loc[nega_bot_ind_ls[i],"OSC"]<df.loc[nega_bot_ind_ls[i+1],"OSC"] and df.loc[nega_bot_ind_ls[i],"波谷最低價"]>df.loc[nega_bot_ind_ls[i+1],"波谷最低價"]:
                        df.loc[nega_bot_ind_ls[i+1],"訊號"] += "底背離＆"
                #TODO 之後要把另一個方向的趨勢背離加上
            df.drop("負向底背離",axis=1,inplace=True)
            
        # 20250324 要先把背離的訊號往後移一天，再判斷止損
        # 這些訊號出現時，是在波峰波谷出現的後一天
        # 所以這裡的訊號要往後移一天
        temp_signal_ls = df["訊號"].tolist()
        temp_signal_ls.insert(0,"")
        df["訊號"] = temp_signal_ls[:-1]
        
        # 判斷止損
        state = ""
        high_stop = 0
        low_stop = 0    
        for i in range(1,len(df)):
            if top_stop_ck:
                if  "頂背離" in df.loc[i,"訊號"]:
                    state = "做空"
                    high_stop = df.loc[i-1,"波峰最高價"] #因為訊號是後一天出現，所以抓前一天波峰最高價
                    continue
                if high_stop and state == "做空":
                    #temp_high = df.loc[i,"High"]
                    temp_close = df.loc[i,"Close"]
                    if temp_close>high_stop:
                        df.loc[i,"訊號"] = "頂背離止損"
                        state = ""
                        high_stop = 0
                        low_stop = 0
            if bot_stop_ck:
                if  "底背離" in df.loc[i,"訊號"]:
                    state = "做多"
                    low_stop = df.loc[i-1,"波谷最低價"] #因為訊號是後一天出現，所以抓前一天波谷最低價
                    continue
                if low_stop and state == "做多":
                    #temp_low = df.loc[i,"Low"]
                    temp_close = df.loc[i,"Close"]
                    if temp_close<low_stop:
                        df.loc[i,"訊號"] = "底背離止損"
                        state = ""
                        high_stop = 0
                        low_stop = 0
        
        #回傳df
        return df

    def decide_position(self,df:pd.DataFrame)->pd.DataFrame:
        """
        依照背離與止損決定部位(做多、做空、平倉)

        Args:
            df (pd.DataFrame): 判斷背離與止損的資料

        Returns:
            df (pd.DataFrame): 判斷部位的資料
        """
        pos = 0
        pos_ls = []
        open_record = []
        for i in df.index.tolist():
            if df.loc[i,"訊號"] == "頂背離止損" or df.loc[i,"訊號"] == "底背離止損":
                if pos != 0:
                    pos = 0
                pos_ls.append(pos)
                open_record.append(True)
            elif "頂背離" in df.loc[i,"訊號"]:
                if pos > 0:
                    pos = -1*pos
                elif pos == 0:
                    pos -= 1
                pos_ls.append(pos)
                open_record.append(True)
            elif "底背離" in df.loc[i,"訊號"]:
                if pos < 0:
                    pos = -1*pos
                elif pos == 0:
                    pos += 1
                pos_ls.append(pos)
                open_record.append(True)
            else:
                pos_ls.append(pos)
                open_record.append(False)
        #因為接觸到訊號隔天才會變化，所以最前面插入原始部位
        pos_ls.insert(0,0)
        pos_ls = pos_ls[:-1]
        open_record.insert(0,False)
        open_record = open_record[:-1]
        #儲存部位
        df["position"] = pos_ls
        df["當時開盤價"] = open_record
        open_record = [df.loc[i,"Open"] if df.loc[i,"當時開盤價"] else pd.NA for i in df.index.tolist()]
        df["當時開盤價"] = open_record
        #回傳df
        return df

    def getIncome(self,df:pd.DataFrame):
        temp_df = df[["日期","時間區間","position","當時開盤價"]]
        temp_df = temp_df.dropna(how="any",axis=0)
        #部位變高->(後Open-前Open)*-1(買回做空的期貨)
        #部位變低->(後Open-前Open)(賣出做多的期貨)
        #做多變做空中間經過平倉，所以是算做多後平倉賣的賺多少
        #做空變做多中間經過平倉，所以是算做空後平倉買回的差價為賺的錢
        income_ls = []
        ind_ls = temp_df.index.tolist()
        for i in range(len(temp_df)):
            if (temp_df.loc[ind_ls[i],"position"]!=0) and (i+1 < len(temp_df)):
                if (temp_df.loc[ind_ls[i],"position"] < temp_df.loc[ind_ls[i+1],"position"]):#部位變高
                    temp_income = (temp_df.loc[ind_ls[i+1],"當時開盤價"]-temp_df.loc[ind_ls[i],"當時開盤價"])*(-1)
                    income_ls.append(temp_income*200)
                elif (temp_df.loc[ind_ls[i],"position"] > temp_df.loc[ind_ls[i+1],"position"]):#部位變低
                    temp_income = (temp_df.loc[ind_ls[i+1],"當時開盤價"]-temp_df.loc[ind_ls[i],"當時開盤價"])
                    income_ls.append(temp_income*200)
                else:
                    income_ls.append(pd.NA)
            else:
                income_ls.append(pd.NA)
                continue
        #income是算在後面那天    
        income_ls.insert(0,pd.NA)
        income_ls = income_ls[:-1]
        temp_df["income($)"] = income_ls
        temp_df = temp_df[["日期","時間區間","income($)"]]
        df = df.merge(temp_df,on=["日期","時間區間"],how="left")
        
        return df

    def getNetIncome(self,df:pd.DataFrame):
        net_income_in_cash = df["income($)"].sum()
        return round(net_income_in_cash)

    def getSumOfCertainCondition(self,val_ls,state:str):
        right_state_ls = []
        for val in val_ls:
            if val is not pd.NA:
                if state == "+" and val>0:
                    right_state_ls.append(val)
                if state == "-" and val<=0:
                    right_state_ls.append(val)
            else:
                continue
        return right_state_ls
    
    def getIncomeRatio(self,df:pd.DataFrame):
        income_ls = df["income($)"].tolist()
        pos_income_ls = self.getSumOfCertainCondition(income_ls,state="+")
        neg_income_ls = self.getSumOfCertainCondition(income_ls,state="-")
        pos_income = sum(pos_income_ls)
        neg_income = sum(neg_income_ls)
        income_ratio = 0
        if neg_income != 0:
            income_ratio = abs(pos_income/neg_income)
        else:
            income_ratio = "-"
            return income_ratio
        return round(income_ratio,2)
    
    def getSignalDetail(self,df:pd.DataFrame):
        signal_data = []
        ind_ls = df.index.tolist()
        for i in ind_ls[:-1]:
            if df.loc[i,"訊號"]:
                next_i = ind_ls[ind_ls.index(i)+1]
                actionStr = ""
                if df.loc[next_i,'position']-df.loc[i,'position'] < 0:
                    actionStr = "賣出"
                elif df.loc[next_i,'position']-df.loc[i,'position'] > 0:
                    actionStr = "買進"
                #交易指令
                signal_content = df.loc[i,"訊號"][:-1] if df.loc[i,"訊號"][-1] == "＆" else df.loc[i,"訊號"]
                signal_data.append([f"{str(datetime.datetime.now()).split(' ')[1][:-3]}","台股指數近月","交易指令",
                                    f"實際部位:{df.loc[i,'position']} 目標部位:{df.loc[next_i,'position']} 價格：市價 (訊號：{signal_content})"])
                #回測成交
                amount = abs(df.loc[next_i,'position']-df.loc[i,'position'])
                if amount != 0:
                    signal_data.append([f"{str(datetime.datetime.now()).split(' ')[1][:-3]}","台股指數近月","回測成交",
                                        f"成交時間:{str(df.loc[next_i,'日期']).split(' ')[0]} {str(df.loc[next_i,'時間區間']).split('~')[0]}({actionStr}) 數量:{amount} 價格:{df.loc[next_i,'當時開盤價']}"])
        signal_df = pd.DataFrame(signal_data,columns=["時間","商品","動作","內容"])
        
        return signal_df

    def getTradeDetail(self,df:pd.DataFrame,principal:float):
        ind_ls = df.index.tolist()
        temp_df = df[["日期","時間區間","position","當時開盤價","income($)"]]
        temp_df = temp_df.dropna(subset=["當時開盤價"],axis=0)
        temp_ind_ls = temp_df.index.tolist()
        trade_data = []
        for i in temp_ind_ls[1:]:
            pre_i = temp_ind_ls[temp_ind_ls.index(i)-1]
            trade_amount = temp_df.loc[i,"position"] - temp_df.loc[pre_i,"position"]
            action1 = ""
            action2 = ""
            if trade_amount<0:
                action1 = "買進"
                action2 = "賣出"
            elif trade_amount>0:
                action1 = "賣出"
                action2 = "買進"
            #act_time_1 = str(df.loc[ind_ls[ind_ls.index(pre_i)-1],"日期"]).split(" ")[0]
            #act_time_2 = str(df.loc[ind_ls[ind_ls.index(i)-1],"日期"]).split(" ")[0]
            #20250322 只要是"當時開盤價"有值的那一行，那行日期就是入場/出場日期
            act_time_date_1 = str(df.loc[ind_ls[ind_ls.index(pre_i)],"日期"]).split(" ")[0]
            act_time_start_1 = str(df.loc[ind_ls[ind_ls.index(pre_i)],"時間區間"]).split("~")[0]
            act_time_date_2 = str(df.loc[ind_ls[ind_ls.index(i)],"日期"]).split(" ")[0]
            act_time_start_2 = str(df.loc[ind_ls[ind_ls.index(i)],"時間區間"]).split("~")[0]
            pre_i_ind_inList = ind_ls.index(pre_i)
            i_ind_inList = ind_ls.index(i)
            
            if temp_df.loc[i,"income($)"] is not pd.NA:
                temp_trade_data = ["台股指數近月(FITXN*1.TF)",f"{i_ind_inList-pre_i_ind_inList}",f"{act_time_date_1} {act_time_start_1}",f"{action1} {temp_df.loc[pre_i,'當時開盤價']}",
                                f"{act_time_date_2} {act_time_start_2}",f"{action2} {temp_df.loc[i,'當時開盤價']}",abs(trade_amount),temp_df.loc[i,"income($)"]]
                trade_data.append(temp_trade_data)
        trade_df = pd.DataFrame(trade_data,columns=["商品名稱","持有區間","進場時間","進場價格","出場時間","出場價格","交易數量","獲利金額"])
        temp_income_ls = []
        income_in_percent_ls = []
        acc_income_ls = []
        acc_income_in_precent_ls = []
        for income in trade_df["獲利金額"]:
            temp_income_ls.append(income)
            acc_income_ls.append(sum(temp_income_ls))
            income_in_percent_ls.append(round((income/principal)*100,2))
            acc_income_in_precent_ls.append(round((sum(temp_income_ls)/principal)*100,2))
        trade_df["累計獲利金額"] = acc_income_ls
        trade_df["報酬率(%)"] = income_in_percent_ls
        trade_df["累計報酬率(%)"] = acc_income_in_precent_ls
        trade_df["序號"] = [i+1 for i in range(len(trade_df))]
        new_cols = ["商品名稱","序號","進場時間","進場價格","出場時間","出場價格","持有區間","報酬率(%)","交易數量","獲利金額","累計獲利金額","累計報酬率(%)"]
        trade_df = trade_df[new_cols]
        return trade_df
    
    def getNetIncomePercent(self,trade_df:pd.DataFrame):
        net_income_in_percent = trade_df["報酬率(%)"].sum()
        return round(net_income_in_percent,2)
    
    def getIncomePercentRatio(self,trade_df:pd.DataFrame):
        income_ls = trade_df["報酬率(%)"].tolist()
        pos_income_ls = self.getSumOfCertainCondition(income_ls,state="+")
        neg_income_ls = self.getSumOfCertainCondition(income_ls,state="-")
        pos_income = sum(pos_income_ls)
        neg_income = sum(neg_income_ls)
        income_perc_ratio = 0
        if neg_income != 0:
            income_perc_ratio = abs(pos_income/neg_income)
        else:
            income_perc_ratio = "-"
            return income_perc_ratio
        return round(income_perc_ratio,2)
    
    def getIntervalDebt(self,trade_df:pd.DataFrame):
        # 先將所有累積損益取出來存成acc_income_ls
        acc_income_ls = trade_df["累計獲利金額"].tolist()
        # 掃描acc_income_ls，如果遇到正數就先取為暫時波峰，
        temp_peak = 0
        debt_ls = []
        for val in acc_income_ls:
            if temp_peak==0:
                if val>0: #取得回測區間第一個正數
                    temp_peak = val
                    continue
                elif val<=0: #負數不管
                    continue
            # 如果後面一個值比波峰大，後面一個值變成新暫時波峰，
            if val >= temp_peak:
                temp_peak = val
            else:# 如果後面一個值比波峰小，儲存區間虧損->[後面一個值-暫時波峰]*-1存到debt_ls
                temp_interval_debt = (val-temp_peak)*-1
                debt_ls.append(temp_interval_debt)
        
        if temp_peak == 0: #如果從頭到尾沒有獲利，回報最大區虧為未知["-"]
            return "-"
        
        if not debt_ls:
            return "-"
        
        # 回傳debt_ls的最大值
        return round(max(debt_ls),2)
    
    def getIntervalPercentDebt(self,trade_df:pd.DataFrame):
        # 先將所有累積損益取出來存成acc_income_ls
        acc_income_ls = trade_df["累計報酬率(%)"].tolist()
        # 掃描acc_income_ls，如果遇到正數就先取為暫時波峰，
        temp_peak = 0
        debt_ls = []
        for val in acc_income_ls:
            if temp_peak==0:
                if val>0: #取得回測區間第一個正數
                    temp_peak = val
                    continue
                elif val<=0: #負數不管
                    continue
            # 如果後面一個值比波峰大，後面一個值變成新暫時波峰，
            if val >= temp_peak:
                temp_peak = val
            else:# 如果後面一個值比波峰小，儲存區間虧損->[後面一個值-暫時波峰]*-1存到debt_ls
                temp_interval_debt = (val-temp_peak)*-1
                debt_ls.append(temp_interval_debt)
        
        if temp_peak == 0: #如果從頭到尾沒有獲利，回報最大區虧為未知["-"]
            return "-"
        
        if not debt_ls:
            return "-"
        
        # 回傳debt_ls的最大值
        return round(max(debt_ls),2)
    
    def checkParam(self):
        Warning_ls = []
        start_date = self.date_val_1.date()
        end_date = self.date_val_2.date()
        now_date = QtCore.QDate.currentDate()
        principal = float(self.principal_val.text()) 
        if start_date >= end_date :
            Warning_ls.append("開始日期須早於結束日期!!\n")
        if start_date>now_date and end_date>now_date:
            Warning_ls.append("開始與結束日期須包含已交易的日期!!\n")
        if not principal:
            Warning_ls.append("請填寫本金!!\n")
        if principal<0:
            Warning_ls.append("本金請填寫正數!!\n")
        
        return Warning_ls
    
    def backTesting(self):
        global df,top_ck,top_stop_ck,bot_ck,bot_stop_ck,positive_top_ck,negative_bot_ck
        df = pd.read_excel('2022-01-01~2025-03-31_小時k_20250401.xlsx') #TODO 之後df需要在回測前更新沒有的資料
        warning_ls = self.checkParam()
        if warning_ls:
            warning_str = ""
            for s in warning_ls: warning_str+=s
            self._info = QtWidgets.QMessageBox(self)
            self._info.information(self,"警告",warning_str)
            return
        start_date = self.date_val_1.date().toString(QtCore.Qt.ISODate)
        end_date = self.date_val_2.date().toString(QtCore.Qt.ISODate)
        data_start_date = self.date_val_1.date().addDays(-3).toString(QtCore.Qt.ISODate)
        #這次不往後抓，因為osc是往前抓資料算出來的
        #data_end_date = self.date_val_2.date().addMonths(3).toString(QtCore.Qt.ISODate)
        data_end_date = end_date #TODO 之後UI的結束日期預設要改成資料最晚日期
        top_ck = self.tacticCheckbtn1.isChecked()
        top_stop_ck = self.tacticCheckbtn2.isChecked()
        bot_ck = self.tacticCheckbtn3.isChecked()
        bot_stop_ck = self.tacticCheckbtn4.isChecked()
        positive_top_ck = self.tacticCheckbtn5.isChecked()
        negative_bot_ck = self.tacticCheckbtn6.isChecked()
        # 這裡需要提醒什麼都沒勾，但沒勾照樣是可以測
        reply = ""
        if  not any([top_ck,top_stop_ck,bot_ck,bot_stop_ck,positive_top_ck,negative_bot_ck]) :
            self._info = QtWidgets.QMessageBox(self)
            reply = self._info.question(self,"提醒","未設定任何策略，確定進行回測?",
                                        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, 
                                        QtWidgets.QMessageBox.No)
        if reply == QtWidgets.QMessageBox.No:
            return
        #小時k資料已經整理為近月、小時k資料
        #df = self.get_TX_data(data_start_date,data_end_date,"TX")
        #df = self.clean_Constract_Date(df)
        #df = self.get_1TF(df)
        #df = self.get_daliy_k_bar_range(df,data_start_date,data_end_date)
        #df = self.getMACD_OSC(df)
        df = self.find_pos(df)
        df = self.read_position(df)
        df = self.decide_position(df)
        df = self.getIncome(df)
        df["訊號"] = [s[:-1] if s and s[-1] == "＆" else s for s in df["訊號"]]
        df = df[(df["日期"]>pd.to_datetime(start_date)) & (df["日期"]<pd.to_datetime(end_date))]
        
        self._info = QtWidgets.QMessageBox(self)
        self._info.information(self,"訊息",f'已完成回測 !!!')
        
        signal_df,trade_df = self.getReportTable(df)
        
        self.RecordTable_1 = QtWidgets.QTableWidget(self)
        self.RecordTable_1.setColumnCount(len(signal_df.columns.tolist()))
        self.RecordTable_1.setHorizontalHeaderLabels(signal_df.columns.tolist())
        columns = signal_df.columns.tolist()
        for r in range(len(signal_df)):
            temp_row = self.RecordTable_1.rowCount()
            self.RecordTable_1.insertRow(temp_row)
            for c in columns:
                tempItem = QtWidgets.QTableWidgetItem(str(signal_df.loc[r,c]))
                self.RecordTable_1.setItem(temp_row,columns.index(c),tempItem)
        #self.putDfToTableUI(signal_df)
        #self.tabwid = QtWidgets.QTabWidget(self)
        self.tab_signal_df_wid = QtWidgets.QWidget(self)
        self.tab_layout_1 = QtWidgets.QVBoxLayout(self.tab_signal_df_wid)
        self.tab_layout_1.addWidget(self.RecordTable_1)
        #self.tabwid.addTab(self.tab_signal_df_wid,"策略交易明細")
        
        self.RecordTable_2 = QtWidgets.QTableWidget(self)
        self.RecordTable_2.setColumnCount(len(trade_df.columns.tolist()))
        self.RecordTable_2.setHorizontalHeaderLabels(trade_df.columns.tolist())
        columns = trade_df.columns.tolist()
        for r in range(len(trade_df)):
            temp_row = self.RecordTable_2.rowCount()
            self.RecordTable_2.insertRow(temp_row)
            for c in columns:
                tempItem = QtWidgets.QTableWidgetItem(str(trade_df.loc[r,c]))
                self.RecordTable_2.setItem(temp_row,columns.index(c),tempItem)
        #self.putDfToTableUI(trade_df)
        #self.tabwid = QtWidgets.QTabWidget(self)
        self.tab_trade_df_wid = QtWidgets.QWidget(self)
        self.tab_layout_2 = QtWidgets.QVBoxLayout(self.tab_trade_df_wid)
        self.tab_layout_2.addWidget(self.RecordTable_2)
        #self.tabwid.addTab(self.tab_trade_df_wid,"買賣交易明細")
        
        self.RecordTable_3 = QtWidgets.QTableWidget(self)
        self.RecordTable_3.setColumnCount(len(df.columns.tolist()))
        self.RecordTable_3.setHorizontalHeaderLabels(df.columns.tolist())
        columns = df.columns.tolist()
        for r in df.index.tolist():
            temp_row = self.RecordTable_3.rowCount()
            self.RecordTable_3.insertRow(temp_row)
            for c in columns:
                tempItem = QtWidgets.QTableWidgetItem(str(df.loc[r,c]))
                self.RecordTable_3.setItem(temp_row,columns.index(c),tempItem)
        #self.putDfToTableUI(trade_df)
        #self.tabwid = QtWidgets.QTabWidget(self)
        self.df_wid = QtWidgets.QWidget(self)
        self.tab_layout_3 = QtWidgets.QVBoxLayout(self.df_wid)
        self.tab_layout_3.addWidget(self.RecordTable_3)
        #self.tabwid.addTab(self.df_wid,"總表")
        
        time_now = datetime.datetime.now()
        
        sample_texts = [f"策略交易明細{str(time_now)[11:19]}",f"買賣交易明細{str(time_now)[11:19]}",f"總表{str(time_now)[11:19]}"]
        for text in sample_texts:
            self.text_list.addItem(text)
        
        # 連接列表的點擊信號
        self.text_list.itemClicked.connect(self.add_new_table_tab)
        self.df_dict[f"策略交易明細{str(time_now)[11:19]}"] = signal_df
        self.df_dict[f"買賣交易明細{str(time_now)[11:19]}"] = trade_df
        self.df_dict[f"總表{str(time_now)[11:19]}"] = df
        #損益歷史加入表格
        cols = ["淨利($)","淨利(%)","平均獲利\n虧損比($)","平均獲利\n虧損比(%)","最大區間虧損($)","最大區間虧損(%)"]
        vals = [self.getNetIncome(df),self.getNetIncomePercent(trade_df),self.getIncomeRatio(df),self.getIncomePercentRatio(trade_df),self.getIntervalDebt(trade_df),self.getIntervalPercentDebt(trade_df)]
        temp_row = self.RecordTable.rowCount()
        self.RecordTable.insertRow(temp_row)
        for i in range(len(cols)):
            tempItem = QtWidgets.QTableWidgetItem(str(vals[i]))
            self.RecordTable.setItem(temp_row,i,tempItem)
        
        if self.his_df.loc[0,"淨利($)"] != "":
            self.his_df.loc[len(self.his_df)] = vals
        else:
            self.his_df = pd.DataFrame(data = [vals],columns=cols)
        self.df_dict["損益歷史紀錄"] = self.his_df
    
    def getReportTable(self,df:pd.DataFrame):
        signal_df = self.getSignalDetail(df)
        principal = float(self.principal_val.text()) 
        trade_df = self.getTradeDetail(df,principal)
        return signal_df,trade_df
    
    def open_selector(self):
        dialog = DataFrameSelector(self.df_dict, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            selected = dialog.get_selected_dataframes()
            self.save_dataframes(selected)
    
    def save_dataframes(self, selected_df_names):
        # 實際儲存DataFrame的邏輯
        saveFolder = QtWidgets.QFileDialog.getExistingDirectory(self,"儲存位置")
        saveFileName,_ok = QtWidgets.QInputDialog.getText(self,"匯出設定","請輸入檔名：")
        if _ok and (not saveFileName):
            self._info = QtWidgets.QMessageBox(self)
            self._info.information(self,"提醒",f'檔名不可空白!!!')
            return 
        savepath = os.path.join(saveFolder,saveFileName+".xlsx")
        if _ok:
            with pd.ExcelWriter(savepath) as wt:
                for df_name in selected_df_names:
                    df = self.df_dict[df_name]
                    temp_sheetName = re.sub(r":","",df_name)
                    df.to_excel(wt,sheet_name=temp_sheetName,index=False)
        else:
            return 
            
        # 顯示儲存成功的訊息
        self._info = QtWidgets.QMessageBox(self)
        self._info.information(self,"訊息",f'已將報表儲存為 {saveFileName}.xlsx !!!')
   
    def putDfToTableUI(self,df:pd.DataFrame):
        columns = df.columns.tolist()
        for r in range(len(df)):
            temp_row = r+1
            self.temp_RecordTable.insertRow(temp_row)
            for c in columns:
                tempItem = QtWidgets.QTableWidgetItem(str(df.loc[r,c]))
                self.temp_RecordTable.setItem(temp_row,columns.index(c),tempItem)
                
    def dataframe_to_table(self, df, table):
        """將 DataFrame 轉換為 QTableWidget"""
        # 設置表格行數和列數
        table.setRowCount(df.shape[0])
        table.setColumnCount(df.shape[1])
        
        # 設置表格的標題行
        table.setHorizontalHeaderLabels(df.columns)
        
        # 填充表格數據
        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iloc[row, col]))
                table.setItem(row, col, item)
        
        return table
    
    def add_new_table_tab(self, item):
        df_name = item.text()
        
        # 檢查是否已經打開了這個標籤頁
        if df_name in self.opened_tabs:
            # 如果已經打開，就切換到該標籤頁
            for i in range(self.tabwid.count()):
                if self.tabwid.tabText(i) == df_name:
                    self.tabwid.setCurrentIndex(i)
                    return
        
        # 獲取相應的 DataFrame
        df = self.df_dict.get(df_name)
        if df is None:
            return
        
        # 創建新的標籤頁
        new_tab = QWidget()
        tab_layout = QVBoxLayout()
        
        # 創建 QTableWidget
        table = QTableWidget()
        
        # 將 DataFrame 轉換為表格
        self.dataframe_to_table(df, table)
        
        # 添加表格到標籤頁
        tab_layout.addWidget(table)
        new_tab.setLayout(tab_layout)
        
        # 添加新標籤頁
        index = self.tabwid.addTab(new_tab, df_name)
        self.tabwid.setCurrentIndex(index)
        
        # 記錄已打開的標籤頁
        self.opened_tabs.append(df_name)
    
    def close_tab(self, index):
        # 獲取標籤頁文字
        tab_text = self.tabwid.tabText(index)
        
        # 移除標籤頁
        self.tabwid.removeTab(index)
        
        # 從已打開列表中移除
        if tab_text in self.opened_tabs:
            self.opened_tabs.remove(tab_text)
    
    def open_top_define_setting(self):
        dialog = AdjustTopReverseDefination_ui(self.top_def_ck1,self.top_def_ck2, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.top_def_ck1,self.top_def_ck2 = dialog.getDefineState()
    
    def open_bot_define_setting(self):
        dialog = AdjustBotReverseDefination_ui(self.bot_def_ck1,self.bot_def_ck2, self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.bot_def_ck1,self.bot_def_ck2 = dialog.getDefineState()
            
class AdjustTopReverseDefination_ui(QtWidgets.QDialog):
    def __init__(self, top_def_ck1, top_def_ck2, parent=None):
        super().__init__(parent)
        #self.parent_widget = parent #可能不需要這一行也可以稱為子視窗
        self.o_top_def_ck1 = top_def_ck1
        self.o_top_def_ck2 = top_def_ck2
        self.setObjectName("TopReverseDefine")
        self.setWindowTitle("頂背離 定義設定")
        self.resize(550,300)
        self.Layout_V = QtWidgets.QVBoxLayout(self)
        
        self.define1_ckb = QtWidgets.QCheckBox(self)
        self.define1_ckb.setText("過去OSC>當前OSC 但 過去最高價<現在最高價")
        self.define1_ckb.setChecked(self.o_top_def_ck1)
        self.Layout_V.addWidget(self.define1_ckb)
        self.define2_ckb = QtWidgets.QCheckBox(self)
        self.define2_ckb.setText("過去OSC<當前OSC 但 過去最高價>現在最高價")
        self.define2_ckb.setChecked(self.o_top_def_ck2)
        self.Layout_V.addWidget(self.define2_ckb)
        
         # 建立按鈕 (確定/取消)
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok,self)
        button_box.accepted.connect(self.accept_check)
        button_box.rejected.connect(self.reject)
        self.Layout_V.addWidget(button_box)
    
    def accept_check(self):
        if (not self.define1_ckb.isChecked()) and (not self.define2_ckb.isChecked()):
            self._info = QtWidgets.QMessageBox(self)
            self._info.information(self,"警告","請至少勾選一項!!!")
            return
        else:
            self.accept()
    
    def getDefineState(self):
        return self.define1_ckb.isChecked(),self.define2_ckb.isChecked()
    
class AdjustBotReverseDefination_ui(QtWidgets.QDialog):
    def __init__(self, bot_def_ck1, bot_def_ck2, parent=None):
        super().__init__(parent)
        #self.parent_widget = parent #可能不需要這一行也可以稱為子視窗
        self.o_bot_def_ck1 = bot_def_ck1
        self.o_bot_def_ck2 = bot_def_ck2
        self.setObjectName("BottomReverseDefine")
        self.setWindowTitle("底背離 定義設定")
        self.resize(550,300)
        self.init_ui()
    
    def init_ui(self):
        self.Layout_V = QtWidgets.QVBoxLayout(self)
        self.define1_ckb = QtWidgets.QCheckBox(self)
        self.define1_ckb.setText("過去OSC<當前OSC 但 過去最低價>現在最低價")
        self.define1_ckb.setChecked(self.o_bot_def_ck1)
        self.Layout_V.addWidget(self.define1_ckb)
        self.define2_ckb = QtWidgets.QCheckBox(self)
        self.define2_ckb.setText("過去OSC>當前OSC 但 過去最低價<現在最低價")
        self.define2_ckb.setChecked(self.o_bot_def_ck2)
        self.Layout_V.addWidget(self.define2_ckb)
        
        # 建立按鈕 (確定/取消)
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok,self)
        button_box.accepted.connect(self.accept_check)
        button_box.rejected.connect(self.reject)
        self.Layout_V.addWidget(button_box)
        # 標記是否已保存設定
        #self.settings_saved = False
    
    def accept_check(self):
        if (not self.define1_ckb.isChecked()) and (not self.define2_ckb.isChecked()):
            self._info = QtWidgets.QMessageBox(self)
            self._info.information(self,"警告","請至少勾選一項!!!")
            return
        else:
            self.accept()
    
    def getDefineState(self):
        return self.define1_ckb.isChecked(),self.define2_ckb.isChecked()
        
class DataFrameSelector(QtWidgets.QDialog):
    def __init__(self, dataframes, parent=None):
        super().__init__(parent)
        self.dataframes = dataframes
        self.selected_dataframes = []
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle('選擇要儲存的報表')
        self.setMinimumWidth(400)
        
        # 創建捲動區域以容納大量的DataFrame
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        scroll_layout = QVBoxLayout(content_widget)
        
        # 顯示說明文字
        info_label = QLabel('請勾選您想要儲存的報表:')
        scroll_layout.addWidget(info_label)
        
        # 建立所有DataFrame的核取方塊
        self.checkboxes = {}
        for df_name in self.dataframes:
            checkbox = QtWidgets.QCheckBox(df_name)
            self.checkboxes[df_name] = checkbox
            scroll_layout.addWidget(checkbox)
        
        scroll.setWidget(content_widget)
        
        # 建立按鈕 (確定/取消)
        button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept_selection)
        button_box.rejected.connect(self.reject)
        
        # 主佈局
        main_layout = QVBoxLayout()
        main_layout.addWidget(scroll)
        main_layout.addWidget(button_box)
        self.setLayout(main_layout)
    
    def accept_selection(self):
        # 收集所有被勾選的DataFrame
        self.selected_dataframes = [df_name for df_name, checkbox in self.checkboxes.items() 
                                  if checkbox.isChecked()]
        
        if not self.selected_dataframes:
            QMessageBox.warning(self, "警告", "您沒有選擇任何報表")
            return
            
        self.accept()
        
    def get_selected_dataframes(self):
        return self.selected_dataframes

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    Form = mainwin()
    Form.show()
    sys.exit(app.exec_())