import sys
sys.path.append('/mnt/d/7.finance_python')
import utility_f as uf
import pandas as pd
import yfinance as yf
import numpy as np
import datetime
import time
import traceback
try:
    data = pd.read_excel('/mnt/d/7.finance_python/stock_list.xlsx', engine='openpyxl') #讀取股票清單
    target_stock = data['代號'].tolist() #讀取股票清單
    today = datetime.date.today() #獲取今日日期
    if_trade = uf.is_open(today) #傳入判斷今天是否為營業日
    if if_trade=='N': #如果是N，表示沒開盤，準備寄信
        mail_list = ['yujay50819@gmail.com'] #沒開盤應該收件者
        subject = f'{today} 小幫手暴跌中的股票偵測 - 今日休市' #標題為三大法人篩選，今日休市
        body = '' #郵件內容為空
        uf.send_mail(mail_list, subject, body, 'text', None, None) #寄信
        exit()
    date_start = today + datetime.timedelta(days =-20) #獲取過去10天的收盤價
    date_start = date_start.strftime('%Y-%m-%d') #轉為str格式，等一下準備轉入yfinance的history獲取歷史股價
    date_end = today + datetime.timedelta(days =1) #獲取T+1日，因為yfinance的end日期是到T+1
    date_end = date_end.strftime('%Y-%m-%d') #轉為str格式，等一下準備轉入yfinance的history獲取歷史股價

    # 創建空list備用
    stock_store = []
    today_store = []
    today_fall = []
    count = 0
    # 迴圈處理每一支目標股票
    for target in target_stock:
        count+=1
        time.sleep(1)
        stock = yf.Ticker(f'{target}.TW') #獲取目標股票的Ticker類
        df = stock.history(start = date_start, end = date_end) #傳入start跟end獲取歷史股價
        # 做一點基本的檢核，防止資料回傳0筆或1筆
        if len(df) >=5:
            today_price = df['Close'].values[-1] #將收盤轉為array並且取最後一筆，今天的收盤
            yes_price = df['Close'].values[-2] #將收盤轉為array並且取倒數第二筆，昨天的收盤
            be_yes_price = df['Close'].values[-3] #將收盤轉為array並且取倒數第三筆，前天的收盤
            
            # [漲跌幅 = (現價 - 上一交易日的收盤價)/上一交易日的收盤價*100%] #
            
            fall_today = ((today_price - yes_price)/yes_price)*100 #根據公式取得今日的漲跌幅
            fall_yes = ((yes_price - be_yes_price)/be_yes_price)*100 #根據公式取得昨日的漲跌幅
            if fall_today <= -5 and fall_yes <= -5:
                print(f'Stock: {target} | fall today: {fall_today} | today: {today_price} | yes: {yes_price}')
                today_store.append(today_price)
                today_fall.append(fall_today)
                stock_store.append(target)
            print(f'Dealing stock: {target} | All stock: {len(target_stock)} | Now: {count}') #print出處理進度

    # 控制用
    print('Get it:', stock_store)
    main_df = pd.DataFrame()
    # loop剛剛篩選的結果
    for t in stock_store:
        time.sleep(1) #使用到跟requests有關的東西，習慣先sleep一下
         #我們要將目標股票的新聞連結再一起，control=0代表第一次loop
        merge_df = uf.get_yahoo_news(t,1) #使用新聞函數獲取新聞，並且先獲取主要的dataframe，其他股票的就連接在後面
        merge_df['stock'] = len(merge_df)*[t] #取得該篇新聞屬於哪一支股票
        main_df = main_df.append(merge_df) #接在最一開始的dataframe後面
    main_df.to_excel(f'/mnt/d/7.finance_python/fall_stock_news.xlsx', index=False) #處理要寄信的部分，首先將剛剛的新聞列表存成excel保存
    # 處理要寄信的部分，這裡處理我希望放在信件body的暴跌中的股票清單
    empty_dataframe = pd.DataFrame()
    empty_dataframe['股票代號'] = stock_store
    empty_dataframe['今價'] = today_store
    empty_dataframe['今日跌幅%'] = today_fall
    # 希望寄出的是表格，因此我們將他轉為html備用
    empty_dataframe = empty_dataframe.to_html(index=False)
    #製作html格式
    body = f'''<html>
                <font face="微軟正黑體"></font>
                <body>
                <h4>
                小幫手系列偵測下表股價為暴跌中股票
                </h4>
                {empty_dataframe}
                <h5>請自行謹慎評估風險</h5>
                </body>
                </html>'''
    # 寄信名單
    mail_list = ['yujay50819@gmail.cpm']
    subject = f'{today} 小幫手暴跌中股票偵測' #標題加上日期，淺顯易懂
    uf.send_mail(mail_list, subject, body,'html', [f'/mnt/d/7.finance_python/fall_stock_news.xlsx'], ['相關新聞表.xlsx'])
except SystemExit:
    print('Its OK')
except:
    today = datetime.date.today()
    mail_list = ['yujay50819@gmail.com']
    subject = f'{today} 小幫手暴跌中股票異常'
    body = traceback.format_exc()
    uf.send_mail(mail_list, subject, body, 'text', None, None)
