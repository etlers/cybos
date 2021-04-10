# -*- coding: utf-8 -*-
import win32com.client
import pandas as pd
import datetime, yaml, os, time
import pymysql
from sqlalchemy import create_engine

# pymysql.install_as_MySQLdb()
# import MySQLdb

# engine = create_engine("mysql+mysqldb://etlers:"+"wndyd"+"@localhost/cybos", encoding='utf-8')
# conn = engine.connect()

# current date and time
now_dtm = datetime.datetime.now()
run_dt = now_dtm.strftime("%Y%m%d")
list_deal_history = []
list_price_info = []
# 설정 값 불러오기
with open('config.yaml', encoding='utf-8') as stream:
    deal_info = yaml.safe_load(stream)
# 거래를 이어갈 파일 명칭
txt_filename = deal_info["txt_filename"]
# 종목코드
jongmok_code = deal_info["jongmok_code"]
samsung_code = deal_info["samsung_code"]
# 최초 지갑
possesion = deal_info["possesion"]
# 전일 종가. 파일에 없는 경우
last_price = deal_info["last_price"]
# 매도 기준이 되는 수익률
profit_rate = deal_info["profit_rate"]
# 매수 기준이 되는 거래 틱 정보
tick_amount = deal_info["tick"]["tick_amount"]
tick_over = deal_info["tick"]["tick_over"]
tick_under = deal_info["tick"]["tick_under"]
# 추출 대기 초
wait_sec = deal_info["wait_sec"]
# 수행시간
start_hms = deal_info["run_hms"]["start_hms"]
stop_hms = deal_info["run_hms"]["stop_hms"]
# 가격정보 컬럼
list_price_info_cols = [
    "DT","JONGMOK_CD","JONGMOK_NM","TM","CPRICE","DIFF","OPEN","HIGH","LOW","OFFER","BID","VOL","VOL_VALUE",
    "EX_PRICE","EX_DIFF","EX_VOL",
]
list_deal_history_cols = [
    "DT","JONGMOK_CD","TM","DIV","PRICE","QTY","PROFIT"
]

# 전 거래내역 저장
def close_data(in_param):
    f = open(txt_filename, "w")
    f.write(in_param)
    f.close()

# 텔레그램 메세지 전송
def send_message(order_div, deal_qty):
    pass

# 최저, 최고 가격 설정
def set_high_low_price(now_price, low_price, high_price):
    # 최초는 그대로 설정
    if (low_price == 0 and high_price == 0):
        low_price = now_price
        high_price = now_price
    # 최초가 아닌경우 처리
    else:
        if low_price > now_price:
            low_price = now_price
        if high_price < now_price:
            high_price = now_price
    
    return low_price, high_price  

def samsung_price():
    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, samsung_code)
    objStockMst.BlockRequest()
    
    # 현재가 통신 및 통신 에러 처리 
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    #print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()
    
    # 현재가 정보 조회
    open= objStockMst.GetHeaderValue(13)  # 시가

    return open

# 현재가 추출
def get_now_price():
    # 현재가를 읽으면서 리스트에 저장. 최종 CSV 파일로 저장
    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, jongmok_code)
    objStockMst.BlockRequest()
    
    # 현재가 통신 및 통신 에러 처리 
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    #print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()
    
    # 현재가 정보 조회
    code = objStockMst.GetHeaderValue(0)  #종목코드
    name= objStockMst.GetHeaderValue(1)  # 종목명
    time= objStockMst.GetHeaderValue(4)  # 시간
    cprice= objStockMst.GetHeaderValue(11) # 종가
    diff= objStockMst.GetHeaderValue(12)  # 대비
    open= objStockMst.GetHeaderValue(13)  # 시가
    high= objStockMst.GetHeaderValue(14)  # 고가
    low= objStockMst.GetHeaderValue(15)   # 저가
    offer = objStockMst.GetHeaderValue(16)  #매도호가
    bid = objStockMst.GetHeaderValue(17)   #매수호가
    vol= objStockMst.GetHeaderValue(18)   #거래량
    vol_value= objStockMst.GetHeaderValue(19)  #거래대금
    
    # 예상 체결관련 정보
    exFlag = objStockMst.GetHeaderValue(58) #예상체결가 구분 플래그
    exPrice = objStockMst.GetHeaderValue(55) #예상체결가
    exDiff = objStockMst.GetHeaderValue(56) #예상체결가 전일대비
    exVol = objStockMst.GetHeaderValue(57) #예상체결수량
    """
    if (exFlag == ord('0')):
        print("장 구분값: 동시호가와 장중 이외의 시간")
    elif (exFlag == ord('1')) :
        print("장 구분값: 동시호가 시간")
    elif (exFlag == ord('2')):
        print("장 구분값: 장중 또는 장종료")
    """
    # 가격정보 저장
    list_temp = []        
    list_temp.append(run_dt)
    list_temp.append(code)
    list_temp.append(name)
    list_temp.append(time)
    list_temp.append(cprice)
    list_temp.append(diff)
    list_temp.append(open)
    list_temp.append(high)
    list_temp.append(low)
    list_temp.append(offer)
    list_temp.append(bid)
    list_temp.append(vol)
    list_temp.append(vol_value)    
    list_temp.append(exPrice)
    list_temp.append(exDiff)
    list_temp.append(exVol)
    # 최종 가격정보 리스트
    list_price_info.append(list_temp)

    return open

# 매수
def order_buy(order_qty):
    pass

# 매도
def order_sell(order_qty):
    pass


# 실제 시작하는 함수
def execute():
    # 초기값 설정
    day_profit = 0
    deal_cnt = 0
    bought_price = 0
    high_price = 0
    low_price = 0
    sell_price = 0    
    end_price = 0
    last_price = 0
    now_price = 0
    price_over = 0
    price_under = 0
    order_div = 1
    # 파일에 존재하는 경우 즉, 어제 매도가 안된 경우 매도를 위한 내역 추출
    try:
        f = open(txt_filename, "r")
        data = f.read()
        last_price = int(data.split(" ")[0])
        order_qty = int(data.split(" ")[1])
        deal_amount = int(data.split(" ")[2])
        f.close()
    # 어제 마지막에 매도까지 한 경우
    except:
        order_qty = 0
        deal_amount = possesion    
    # 전일 매도 못한 경우로 매도로 설정
    if order_qty > 0:
        order_div = 2
        sell_price = last_price
        bought_price = last_price      
    
    # 지정한 시간 동안 수행
    while True:
        now = datetime.datetime.now() # current date and time
        run_hms = now.strftime("%H%M%S")
        run_hms_split = now.strftime("%H:%M:%S")
        # 시작 전이면 대기 메세지 출력하면서 대기
        if run_hms < start_hms:
            print("Waiting...", run_hms)
            time.sleep(1)
            continue
        # 시간이 지났으면 종료하러 나감
        if run_hms > stop_hms: break
        # 체결금액 추출
        samsung = samsung_price()
        #comp_price = int(samsung * 0.31)
        now_price = get_now_price()
        comp_rt = round(round(now_price / samsung, 2) * 100, 2)
        print(run_hms_split, now_price, samsung, comp_rt)
        # 최고, 최저 가격 설정
        # 마감 15분 전부터 매도인 경우 최고가보다 크면 매도 매수인 경우는 최저가보다 작으면 매수
        #low_price, high_price = set_high_low_price(now_price, low_price, high_price)
        # 매도까지 끝난 경우의 종가
        # 현재가로 매수, 매도 확인
        # 매도까지 끝난 경우의 종가로 사용
        end_price = now_price
        list_deal = []
        # 매수
        if order_div == 1:
            # 매도 마지막 금액으로 매수할 구간 설정
            price_over = last_price + (tick_amount * tick_over)
            price_under = last_price - (tick_amount * tick_under)
            # 매수 구간에 들어온 경우
            if (now_price > price_over or now_price < price_under):
                order_qty = int(deal_amount / now_price)
                order_buy(order_qty)
                list_deal.append(run_dt)
                list_deal.append(run_hms)
                list_deal.append("BUY")
                list_deal.append(now_price)
                list_deal.append(order_qty)
                list_deal.append(0)
                deal_cnt += 1
                order_div = 2
                bought_price = now_price
                sell_price = int(now_price + now_price * profit_rate)
        # 매도
        else:
            if now_price > sell_price:
                day_profit = day_profit + ((now_price * order_qty) - (bought_price * order_qty))
                order_sell(order_qty)
                list_deal.append(run_dt)
                list_deal.append(run_hms)
                list_deal.append("SELL")
                list_deal.append(now_price)
                list_deal.append(order_qty)
                list_deal.append(day_profit)
                deal_cnt += 1
                order_div = 1
                last_price = bought_price
                order_qty = 0
        # 거래내역이 존재하는 경우      
        if len(list_deal) > 0:
            list_deal_history.append(list_deal)
        # 대기
        time.sleep(wait_sec)

    # 당일 매도 못한 경우 파일에 저장하여 다음 날에 매도 진행  
    if bought_price > 0:
        close_data(str(bought_price) + " " + str(order_qty) + " " + str(deal_amount + day_profit))
    # 당일 매도까지 했다면 최종 가격을 종가로
    else:
        close_data(str(end_price) + " " + "0" + " " + str(deal_amount + day_profit))
    # 일 거래내역 회신
    return day_profit, deal_cnt, deal_amount

# Start Deal
if __name__ == "__main__":
    # 연결
    instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    if instCpCybos.IsConnect == 1:
        print("Connected...")
    else:
        print("Not connected")
    # 실제 프로세스
    execute()
    # 헤더 출력
    print("#" * 50)
    print("Day ", "Profit ", "deal ", " Rate", " Deal")
    print("#" * 50)
    # 최종 결과
    profit, deal_cnt, deal_amount = execute()
    print(run_dt, profit, deal_amount, str(round(round(profit / possesion, 4) * 100, 2)) + "%", deal_cnt)
    # 디비로 저장
    df_price_info = pd.DataFrame(list_price_info, columns=list_price_info_cols)
    df_price_info.to_csv("./csv/price_info.csv", index=False)
    # df_price_info.to_sql(name="price_info", con=engine, if_exists='append', index=False)
    df_deal_history = pd.DataFrame(list_deal_history, columns=list_deal_history_cols)
    df_deal_history.to_csv("./csv/deal_history.csv", index=False)
    # df_deal_history.to_sql(name="deal_history", con=engine, if_exists='append', index=False)