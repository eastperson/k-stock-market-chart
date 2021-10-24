import requests
import pandas as pd
import win32com.client
import time
import schedule

class Stock:
    code = ''
    name = ''
    times = ''
    cprice = ''
    diff = ''
    open = ''
    high = ''
    low = ''
    offer = ''
    bid = ''
    vol = ''
    vol_value = ''
    exFlag = ''
    exPrice = ''
    exDiff = ''
    exVol = ''

    def __init__(self, code, name, times, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol):
        self.code = ''
        self.name = ''
        self.times = ''
        self.cprice = ''
        self.diff = ''
        self.open = ''
        self.high = ''
        self.low = ''
        self.offer = ''
        self.bid = ''
        self.vol = ''
        self.vol_value = ''
        self.exFlag = ''
        self.exPrice = ''
        self.exDiff = ''
        self.exVol = ''

    def __str__(self):
        return "code : {}, name : {}, time : {}, cprice : {}, diff : {}, open : {}, high : {}, low : {}, " \
               "offer : {}, bid : {}, vol : {}, vol_value : {}, exFlag : {}, exPrice : {}, exDiff : " \
               "{}, exVol : {}" \
            .format(self.code, self.name, self.times, self.cprice, self.diff, self.open, self.high,
                    self.low,
                    self.offer, self.bid, self.vol, self.vol_value, self.exFlag, self.exPrice,
                    self.exDiff, self.exVol)

def daily_stock():
    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()

    # 종목코드 리스트 구하기
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1) #거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2) #코스닥

    # 현재가 객체 구하기
    objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
    objStockMst.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
    objStockMst.BlockRequest() 
    
    # 현재가 통신 및 통신 에러 처리 
    rqStatus = objStockMst.GetDibStatus()
    rqRet = objStockMst.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        exit()

    stockList = []

    print("거래소 종목코드", len(codeList))
    print("순서", "종목코드","종목명","시간","종가","대비","시가","고가","저가","매도호가","매수호가","거래량","거래대금"
    ,"예상체결가 구분 플래그","예상체결가","예상체결가 전일대비","예상체결수량")
    cnt = 0
    for i, code in enumerate(codeList):
        objStockMst.SetInputValue(0, code)
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리 
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)  #종목코드
        name= objStockMst.GetHeaderValue(1)  # 종목명
        times= objStockMst.GetHeaderValue(4)  # 시간
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
        
        cnt += 1
        if cnt == 50 :
            cnt = 0
            time.sleep(15)
        print(i, code, name, times, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol)
        stock = Stock(code, name, times, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol)
        result = []
        result.append(code)
        result.append(name)
        result.append(times)
        result.append(cprice)
        result.append(diff)
        result.append(open)
        result.append(high)
        result.append(low)
        result.append(offer)
        result.append(bid)
        result.append(vol)
        result.append(vol_value)
        result.append(exFlag)
        result.append(exPrice)
        result.append(exDiff)
        result.append(exVol)
        stockList.append(result)

    for i, code in enumerate(codeList2):
        objStockMst.SetInputValue(0, code)
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리 
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)  #종목코드
        name= objStockMst.GetHeaderValue(1)  # 종목명
        times= objStockMst.GetHeaderValue(4)  # 시간
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
        
        cnt += 1
        if cnt == 50 :
            cnt = 0
            time.sleep(15)
        print(i + len(codeList), code, name, times, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol)
        stock = Stock(code, name, times, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol)
        result = []
        result.append(code)
        result.append(name)
        result.append(times)
        result.append(cprice)
        result.append(diff)
        result.append(open)
        result.append(high)
        result.append(low)
        result.append(offer)
        result.append(bid)
        result.append(vol)
        result.append(vol_value)
        result.append(exFlag)
        result.append(exPrice)
        result.append(exDiff)
        result.append(exVol)
        stockList.append(result)

    try:
        print(stockList)
        path = "C:\\Users\\kjuio\\file_test"
        stock_df = pd.DataFrame(stockList,
                                columns=["종목코드","종목명","시간","종가","대비","시가","고가","저가","매도호가","매수호가","거래량","거래대금"
                                    ,"예상체결가 구분 플래그","예상체결가","예상체결가 전일대비","예상체결수량"])
        stock_df.index = stock_df.index + 1
        tm = time.localtime()
        string = time.strftime('%Y_%m_%d', tm)
        title = '_'+string+'.csv'
        stock_df.to_csv(path,'stock'+title, mode='w', encoding='utf-8-sig', header=True, index=True)
    except PermissionError:
        path = "C:\\Users\\kjuio\\file_test"
        stock_df = pd.DataFrame(stockList,
                                columns=["순서", "종목코드","종목명","시간","종가","대비","시가","고가","저가","매도호가","매수호가","거래량","거래대금"
                                    ,"예상체결가 구분 플래그","예상체결가","예상체결가 전일대비","예상체결수량"])
        stock_df.index = stock_df.index + 1
        stock_df.to_csv(path,'stock'+title, mode='w', encoding='utf-8-sig', header=True, index=True)
# -9 시간을 해줘야 한다.
schedule.every().day.at("03:30").do(daily_stock)
print('start')
while True:
    schedule.run_pending()          
    time.sleep(1)
