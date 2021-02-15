import pandas as pd
import numpy as np
from PyQt5.QAxContainer import *  # 응용프로그램 제어용
from PyQt5.QtCore import *
from autostock.config.errorCode import *
from PyQt5.QtTest import *

# QTest.qWait(3600) ## event loop에서 3.6초간 process를 기다려 준다.


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom Class입니다.')
        #self.logging.logger.debug("Kiwoom() class start.")
        ##### event loop 모음 ############################
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.detail_account_info_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.calculator_event_loop = QEventLoop()
        ##################################################

        ##### 요청 스크린번호(화면번호) 모음 ############################
        self.screen_my_info = "2000"  # 계좌 관련한 스크린 번호
        self.screen_calculation_stock = "4000"  # 계산용 스크린 번호
        self.screen_real_stock = "5000"  # 종목별 할당할 스크린 번호
        self.screen_meme_stock = "6000"  # 종목별 할당할 주문용 스크린 번호
        self.screen_start_stop_real = "1000"  # 장 시작/종료 실시간 스크린 번호

        ##################################################

        ##### 변수 모음 ############################
        self.account_num = None
        self.account_stock_dict = {}  # 계좌평가잔고내역요청에 따른 계좌내 종목 dict
        self.not_account_stock_dict = {}  # 미체결내역 dict

        ##################################################

        ##### 계좌 거래 관련 변수 모음 ############################
        self.use_money = 0
        self.use_money_percent = 0.5  # 계좌에서 거래에 사용할 금액 비중.

        ##################################################

        ### 초기 셋팅 함수들 바로 실행 ###############################
        self.get_ocx_instance()  # Kiwoom OpenKPI 접속, OCX 방식을 파이썬에 사용할 수 있게 반환
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음(이벤트 처리구역과 tr목록구역)
        self.signal_login_commConnect()  # 로그인 요청 함수 포함
        self.get_account_info()     # 계좌번호 가져오기
        self.detail_account_info()  # 예수금 요청
        self.detail_account_mystock()  # 계좌평가잔고내역 요청
        # 5초 뒤에 미체결 종목들 가져오기 실행
        QTimer.singleShot(5000, self.not_concluded_account)

        self.calculator_fnc()  # 종목분석 실행 (임시용)

        ##################################################

    def get_ocx_instance(self):  # Kiwoom OpenKPI 접속
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')  # 레지스트리에 등록된 이름

    def event_slots(self):  # 이벤트 처리구역과 tr목록구역 생성
        self.OnEventConnect.connect(self.login_slot)  # 로그인 처리 진행용.
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr 목록 data 가져오는 용도
        # self.OnReceiveMsg.connect(self.msg_slot)

    # 로그인은 CommConnect()함수를 호출하며 OnEventConnect 이벤트 인자값으로 로그인

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널

        self.login_event_loop.exec_()  # 이벤트 루프 실행

    def login_slot(self, errCode):
        print(errors(errCode))
        # self.logging.logger.debug(errors(err_code)[1])
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall(
            "GetLoginInfo(QString)", "ACCNO")  # 계좌번호 반환
        account_num = account_list.split(';')[0]  # a;b;c  [a, b, c]

        self.account_num = account_num
        print('나의 보유계좌번호: %s' % self.account_num)  # 8158893311
        #self.logging.logger.debug("계좌번호 : %s" % account_num)

    def detail_account_info(self, sPrevNext="0"):  # 예수금 요청 부분

        self.dynamicCall('SetInputValue(QString, QString)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(QString, QString)', '조회구분', '2')
        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '예수금상세현황요청', 'opw00001',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

        # Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # CommRqData( "RQName"	,  "opw00001"	,  "0"	,  "화면번호"); 위의 2000번
        # 화면번호(=스크린번호)는 grouping을 화면번호당 100개 저장 가능
        # 화면번호는 200개까지 만들수 있음.
        # 한페이지 20개 종목까지검색가능 sPrevNext =0 다음페이지없음. 2 다음페이지

    def detail_account_mystock(self, sPrevNext='0'):  # 계좌평가잔고내역요청
        print('계좌평가잔고내역요청_연속조회: %s' % sPrevNext)

        self.dynamicCall('SetInputValue(QString, QString)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(QString, QString)', '조회구분', '2')
        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '계좌평가잔고내역요청', 'opw00018',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext='0'):  # 실시간미체결현황요청
        print('실시간미체결현황요청')
        #self.logging.logger.debug("미체결 종목 요청")
        self.dynamicCall('SetInputValue(QString, QString)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)', '매매구분', '0')
        self.dynamicCall('SetInputValue(QString, QString)', '체결구분', '1')
        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '실시간미체결현황', 'opt10075', sPrevNext, self.screen_my_info)

########################## 요청에 대해서 받는 내용들 - TR 목록 #######################

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == '예수금상세현황요청':  # TR 목록의 opw0001에 있는 output 항목을 가져올 수 있음.
            deposit = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '예수금')

            self.deposit = int(deposit)
            use_money = float(self.deposit) * self.use_money_percent
            self.use_money = int(use_money)
            # self.use_money = self.use_money / 4

            max_order_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '주문가능금액')

            print('예수금: %s\n주문가능금액: %s' %
                  (int(deposit), int(max_order_amount)))
            #self.logging.logger.debug('예수금: %s\n주문가능금액: %s' % (int(deposit), int(max_order_amount)))
            # self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

        # TR 목록의 opw00018에 있는 output 항목을 가져올 수 있음.
        elif sRQName == '계좌평가잔고내역요청':
            total_buy_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총매입금액')
            self.total_buy_amount = int(total_buy_amount)

            total_profit_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총평가손익금액')
            self.total_profit_amount = int(total_profit_amount)

            total_profit_rate = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총수익률(%)')
            self.total_profit_rate = float(total_profit_rate)

            print('총매입금액: %s\n총평가손익금액: %s\n총수익률(%%): %s' % (
                total_buy_amount, total_profit_amount, total_profit_rate))
            #self.logging.logger.debug("계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (total_buy_amount, total_profit_amount, total_profit_rate))

            rows = self.dynamicCall(
                'GetRepeatCnt(QString, QString', sTrCode, sRQName)  # 한page에 20개까지만 불러오기 가능.

            for i in range(rows):
                code = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목번호')
                code_nm = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
                learn_rate = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '수익률(%)')
                buy_price = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '매입가')
                current_price = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                stock_quantity = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '보유수량')
                total_maeip_amount = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '매입금액')
                current_amount = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '평가금액')

                #self.logging.logger.debug("종목번호: %s - 종목명: %s - 수익률: %s - 매입가:%s - 현재가: %s - 보유수량: %s " % (code, code_nm, learn_rate, buy_price, current_price, stock_quantity))

                code = code.strip()[1:]  # 공란을 제외하고, 두번자 글자부터 마지막까지
                code_nm = code_nm.strip()
                learn_rate = float(learn_rate.strip())
                buy_price = int(buy_price.strip())
                current_price = int(current_price.strip())
                stock_quantity = int(stock_quantity.strip())
                possible_quantity = int(possible_quantity.strip())
                total_maeip_amount = int(total_maeip_amount.strip())
                current_amount = int(current_amount.strip())

                if code in self.account_stock_dict:
                    pass

                else:
                    self.account_stock_dict[code] = {}

                asd = self.account_stock_dict[code]

                asd.update({'종목명': code_nm})
                asd.update({'수익률(%)': learn_rate})
                asd.update({'매입가': buy_price})
                asd.update({'현재가': current_price})
                asd.update({'보유수량': stock_quantity})
                asd.update({'매매가능수량': possible_quantity})
                asd.update({'매입금액': total_maeip_amount})
                asd.update({'평가금액': current_amount})

            print('내 계좌에 있는 종목수: %s' % len(self.account_stock_dict))
            #self.logging.logger.debug("sPreNext : %s" % sPrevNext)
            #self.logging.logger.debug("계좌에 가지고 있는 종목은 %s " % rows)

            if sPrevNext == '2':
                self.detail_account_mystock(sPrevNext='2')
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == '실시간미체결현황':
            rows = self.dynamicCall(
                'GetRepeatCnt(QString, QString', sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목번호')
                code_nm = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
                order_no = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '주문번호')
                order_status = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '주문상태')  # 접수, 확인, 체결
                order_quantity = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '주문수량')
                order_price = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '주문가격')
                current_price = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                order_gubun = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '주문구분')  # -매도, +매수, 정정, 취소
                not_quantity = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '미체결수량')
                ok_quantity = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '체결량')

                code = code.strip()  # code 첫글자에 영문이 없음. 그대로 사용
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                current_price = int(current_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')  # 매수, 매도 앞에 +,- 를 지움
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                nasd = self.not_account_stock_dict[order_no]

                nasd.update({'종목코드': code})
                nasd.update({'종목명': code_nm})
                nasd.update({'주문번호': order_no})
                nasd.update({'주문상태': order_status})
                nasd.update({'주문수량': order_quantity})
                nasd.update({'주문가격': order_price})
                nasd.update({'현재가': current_price})
                nasd.update({'주문구분': order_gubun})
                nasd.update({'미체결수량': not_quantity})
                nasd.update({'체결량': ok_quantity})

                print('미체결종목: %s' % self.not_account_stock_dict[order_no])
                #self.logging.logger.debug("미체결 종목 : %s " % self.not_account_stock_dict[order_no])

            self.detail_account_info_event_loop.exit()

        elif sRQName == "주식일봉차트조회":

            code = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
            # print(code)
            code = code.strip()
            # data = self.dynamicCall("GetCommDataEx(QString, QString)", sTrCode, sRQName)

            print('%s 주식일봉데이터요청' % code)

            cnt = self.dynamicCall(
                "GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            print(cnt)
            # self.logging.logger.debug("남은 일자 수 %s" % cnt)

            if sPrevNext == '2':
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)

            else:
                self.calculator_event_loop.exit()

            self.detail_account_info_event_loop.exit()

#### 주식 선택용 #############################################################

    def get_code_list_by_market(self, market_code):  # 종목코드들 반환
        code_list = self.dynamicCall(
            'GetCodeListByMarket(QString)', market_code)
        code_list = code_list.split(';')[:-1]
        print(code_list)
        return code_list

    # def get_code_list_by_mystock(self, mystock_code):  # 관심종목코드를 불러와서 사용
    #     pd.read_excel()
    #     return my_code_list

    def calculator_fnc(self):  # 종목분석 실행용 함수
        # code_list = self.get_code_list_by_market('0') #코스피 마켓
        code_list = self.get_code_list_by_market('10')  # 코스닥 마켓
        print('코스닥 종목 갯수: %s' % len(code_list))
        #self.logging.logger.debug("코스닥 종목 갯수: %s " % len(code_list))

        # code_list 안에서 index와 value를 모두 사용.enumerate
        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)",
                             self.screen_calculation_stock)  # 스크린 연결 끊기
            print('%s / %s : KOSDAQ Stock Code : %s is updating... ' %
                  (idx + 1, len(code_list), code))
            #self.logging.logger.debug("%s / %s : KOSDAQ Stock Code : %s is updating... " % (idx + 1, len(code_list), code))

            self.day_kiwoom_db(code=code)

    def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):  # 주식일봉차트조회
        QTest.qWait(3600)  # 3.6초마다 딜레이를 준다..

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "주식일봉차트조회", "opt10081", sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec_()


# account_num = 8158893311


#        Candidates are:
# 1 / 1481 : KOSDAQ Stock Code : 900080 is updating...
# QAxBase::dynamicCallHelper: DiscunnectRealData(QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# 2 / 1481 : KOSDAQ Stock Code : 900110 is updating...
# QAxBase::dynamicCallHelper: GetCommDate(QString,QString,int,QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:

# <class 'NoneType'>
# None 주식 일봉데이터요청
# 600
# QAxBase: Error calling IDispatch member SetInputValue: Non-optional parameter missing
# QAxBase::dynamicCallHelper: DiscunnectRealData(QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# 3 / 1481 : KOSDAQ Stock Code : 900270 is updating...
# QAxBase::dynamicCallHelper: GetCommDate(QString,QString,int,QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# <class 'NoneType'>
# None 주식 일봉데이터요청
# 0
# QAxBase::dynamicCallHelper: GetCommDate(QString,QString,int,QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# <class 'NoneType'>
# None 주식 일봉데이터요청
# 0
# QAxBase: Error calling IDispatch member SetInputValue: Non-optional parameter missing
# QAxBase::dynamicCallHelper: DiscunnectRealData(QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# 4 / 1481 : KOSDAQ Stock Code : 900260 is updating...
# QAxBase::dynamicCallHelper: GetCommDate(QString,QString,int,QString): No such property in {a1574a0d-6bfa-4bd7-9020-ded88711818d} [KHOpenAPI Control]
#         Candidates are:
# <class 'NoneType'>
# None 주식 일봉데이터요청
