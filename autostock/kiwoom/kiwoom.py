import pandas as pd
import numpy as np
import os
from PyQt5.QAxContainer import *  # 응용프로그램 제어용
from PyQt5.QtCore import *
from autostock.config.errorCode import *
from PyQt5.QtTest import *

# QTest.qWait(3600) ## event loop에서 3.6초간 process를 기다려 준다.


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('\n=====Kiwoom Open API Start======\n')
        #self.logging.logger.debug("Kiwoom() class start.")
        ##### event loop 모음 ############################
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.detail_account_info_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.calculator_event_loop = QEventLoop()
        ##################################################

        ##### 요청 스크린번호(화면번호) 모음 ############################
        self.screen_my_info = "2000"  # 계좌 관련한 스크린 번호
        self.screen_calculation_stock = "4000"  # 계산용 스크린 번호
        # self.screen_real_stock = "5000"  # 종목별 할당할 스크린 번호
        # self.screen_meme_stock = "6000"  # 종목별 할당할 주문용 스크린 번호
        # self.screen_start_stop_real = "1000"  # 장 시작/종료 실시간 스크린 번호

        ##################################################

        ##### 변수 모음 ############################
        self.dir = 'C:/Users/histi/py37_32/autostock/'
        # max_day = 20 max_day 이상 고가/저가가 이평선 위에 있는 경우. 변수 검색하여 조정
        self.account_num = None
        self.account_stock_dict = {}  # 계좌평가잔고내역요청에 따른 계좌내 종목 dict
        self.not_account_stock_dict = {}  # 미체결내역 dict

        ##################################################

        ##### 종목분석용 변수 모음 ##########################
        self.calcul_data = []

        ##### 계좌 거래 관련 변수 모음 ############################
        self.use_money = 0  # 실제 투자에 사용할 금액
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
        # QTimer.singleShot(5000, self.not_concluded_account)

        # self.calculator_fnc()  # 종목분석 실행 (임시용)
        self.read_code()  # 저장된 종목을 불러온다.

        ##################################################

        # QTest.qWait(10000)
        # self.read_code()
        # self.screen_number_setting()

        # QTest.qWait(5000)

        # # 실시간 수신 관련 함수
        # self.dynamicCall("SetRealReg(QString, QString, QString, QString)",
        #                  self.screen_start_stop_real, '', self.realType.REALTYPE['장시작시간']['장운영구분'], "0")

        # for code in self.portfolio_stock_dict.keys():
        #     screen_num = self.portfolio_stock_dict[code]['스크린번호']
        #     fids = self.realType.REALTYPE['주식체결']['체결시간']
        #     self.dynamicCall(
        #         "SetRealReg(QString, QString, QString, QString)", screen_num, code, fids, "1")

        # self.slack.notification(
        #     pretext="주식자동화 프로그램 동작",
        #     title="주식 자동화 프로그램 동작",
        #     fallback="주식 자동화 프로그램 동작",
        #     text="주식 자동화 프로그램이 동작 되었습니다."
        # )

    def get_ocx_instance(self):  # Kiwoom OpenKPI 접속
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')  # 레지스트리에 등록된 이름

    def event_slots(self):  # 이벤트 처리구역과 tr목록구역 생성
        self.OnEventConnect.connect(self.login_slot)  # 로그인 처리 이벤트 용도.
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr(트랜잭션) 요청관련 이벤트 용도.
        # self.OnReceiveMsg.connect(self.msg_slot)

    # 로그인은 CommConnect()함수를 호출하며 OnEventConnect 이벤트 인자값으로 로그인

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널

        self.login_event_loop.exec_()  # 이벤트 루프 실행

    def login_slot(self, err_code):
        print(errors(err_code))
        # self.logging.logger.debug(errors(err_code)[1])
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall(
            "GetLoginInfo(QString)", "ACCNO")  # 계좌번호 반환
        account_num = account_list.split(';')[0]  # a;b;c   [a, b, c]

        self.account_num = account_num
        print('\n보유계좌번호: %s' % self.account_num)  # 8158893311
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
        print('계좌평가잔고내역요청_연속조회: %s\n' % sPrevNext)

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

            print('예수금: %s\n주문가능금액: %s\n' %
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

            print('총매입금액: %s\n총평가손익금액: %s\n총수익률(%%): %s\n' % (
                int(total_buy_amount), int(total_profit_amount), float(total_profit_rate)))
            #self.logging.logger.debug("계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (int(total_buy_amount), int(total_profit_amount), float(total_profit_rate)))

            # 한page에 20개까지만 불러오기 가능.

            rows = self.dynamicCall(
                "GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            print('내 계좌에 있는 종목수: %s\n' % rows)

            for i in range(rows):
                code = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '종목번호')
                code_nm = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '종목명')
                learn_rate = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '수익률(%)')
                buy_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '매입가')
                current_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '현재가')
                stock_quantity = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '보유수량')
                total_maeip_amount = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '매입금액')
                current_amount = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '평가금액')

                #self.logging.logger.debug("종목번호: %s - 종목명: %s - 수익률: %s - 매입가:%s - 현재가: %s - 보유수량: %s " % (code, code_nm, learn_rate, buy_price, current_price, stock_quantity))

                code = code.strip()[1:]  # 공란을 제외하고, 두번자 글자부터 마지막까지
                code_nm = code_nm.strip()
                learn_rate = float(learn_rate.strip())
                buy_price = int(buy_price.strip())
                current_price = int(current_price.strip())
                stock_quantity = int(stock_quantity.strip())
                total_maeip_amount = int(total_maeip_amount.strip())
                current_amount = int(current_amount.strip())

                print('종목명: %s\n수익률(%%): %s\n매입가: %s\n현재가: %s\n보유수량: %s\n매입금액: %s\n평가금액: %s\n' % (
                    code_nm, learn_rate, buy_price, current_price, stock_quantity, total_maeip_amount, current_amount))

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
                asd.update({'매입금액': total_maeip_amount})
                asd.update({'평가금액': current_amount})

            print('내 계좌에 있는 종목수: %s\n' % len(self.account_stock_dict))
            #self.logging.logger.debug("sPrevNext : %s" % sPrevNext)
            #self.logging.logger.debug("계좌에 가지고 있는 종목은 %s " % rows)

            if sPrevNext == '2':
                self.detail_account_mystock(sPrevNext='2')
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == '실시간미체결현황':
            rows = self.dynamicCall(
                "GetRepeatCnt(QString, QString)", sTrCode, sRQName)
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

                print('미체결종목: %s\n' % self.not_account_stock_dict[order_no])
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

            print('데이터 일수: %s\n' % cnt)
            # self.logging.logger.debug("남은 일자 수 %s" % cnt)

            # 한번조회할때 600일치까지 일봉 data를 받을 수 있다.
            for i in range(cnt):
                # data = self.dynamicCall('GetCommData(QString, QString)', sTrCode, sRQName)
                # 결과물 [['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가',''],['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가','']]
                data = []

                current_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                value = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래량")
                trading_value = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "거래대금")
                date = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "일자")
                start_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "시가")
                high_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가")
                low_price = self.dynamicCall(
                    "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가")

                data.append("")
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append("")

                self.calcul_data.append(data.copy())
                # calcul_data 내용 [['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가',''], [....]]

            if sPrevNext == '2':
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)

            else:
                print('총 일수 %s' % len(self.calcul_data))

                ### 종목 선정 조건 설정 ################################
                # 1. 그랜빌의 매수신호 4법칙 계산 ######################

                pass_success = False

                # 120일 이평선 조건 사용
                if self.calcul_data == None or len(self.calcul_data) < 120:
                    pass_success = False

                else:
                    # 120일 이평선의 최근 가격 구함
                    total_price = 0
                    for value in self.calcul_data[:120]:
                        total_price += int(value[1])
                    moving_average_price = total_price / 120

                    # 오늘자 주가가 120일 이평선에 걸쳐 있는지 확인
                    # 오늘의 저가가 120일 이평선 아래이고 고가는 120일 이평선 위에 있는경우
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(self.calcul_data[0][6]):
                        print('오늘 주가 120일 이평선에 걸쳐 있는지 확인')
                        #self.logging.logger.debug("오늘의 주가가 120 이평선에 걸쳐있는지 확인")
                        bottom_stock_price = True
                        check_price = int(self.calcul_data[0][6])  # 고가 확인

                    # 과거 주가가 계속 120일 이평선보다 밑에 있다가 올라왔는지 확인
                    low_price_prev = None  # 과거 일봉 저가
                    if bottom_stock_price == True:
                        moving_average_price_prev = 0  # 과거의 120일 이평선
                        price_top_moving = False  # 이평선 위로 가 있으면 True

                        idx = 1
                        while True:

                            # 어제 이후 120일치 있는지 확인
                            if len(self.calcul_data[idx:]) < 120:
                                print('120일치 data 없음')
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price / 120

                            # max_day이상 고가가 120일 이평선 위에 있을 경우 제외
                            max_day = 30
                            if int(self.calcul_data[idx][6]) >= moving_average_price_prev and idx <= max_day:
                                print('### 고가가 %s일 이상 120일 이평선 위에 있어서 제외됨' %
                                      max_day)
                                price_top_moving = False
                                break

                            # 저가가 120일 이평선 위에 있었다면 제외
                            elif int(self.calcul_data[idx][7]) > moving_average_price_prev and idx > max_day:
                                print('### 저가가 %s일 이상 120일 이평선 위에 있어서 제외됨' %
                                      max_day)
                                # self.logging.logger.debug('### 저가가 %s일 이상 120일 이평선 위에 있어서 제외됨' % max_day)
                                price_top_moving = True  # 저가가 120일선 이평선 위에 존재
                                # 이평선 위에 존재하는 저가를 저장
                                low_price_prev = int(self.calcul_data[idx][7])
                                break

                            idx += 1

                        # 해당 부분 이평선이 가장 최근 일자의 이평선 가격보다 낮은지 확인
                        if price_top_moving == True:
                            if moving_average_price > moving_average_price_prev and check_price > low_price_prev:
                                print(
                                    '### 포착된 이평선 가격이 오늘자 이평선 가격보다 낮고\n### 과거 일봉 저가가 오늘자 일봉의 고가보다 낮은 것 확인됨')
                                #self.logging.logger.debug("포착된 이평선의 가격이 오늘자 이평선 가격보다 낮은 것 확인")
                                #self.logging.logger.debug("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확인")
                                pass_success = True

                if pass_success == True:
                    print('조건부 통과\n')

                    code_nm = self.dynamicCall(
                        'GetMasterCodeName(QString)', code)

                    f = open('%sfiles/condition_stock.txt' %
                             dir, 'a', encoding='utf8')  # a는 연결해서 쓴다. w는 덮어쓴다.
                    f.write('%s\t%s\t%s\n' %
                            (code, code_nm, str(self.calcul_data[0][1])))
                    f.close()

                elif pass_success == False:
                    print('조건부 통과 못함\n')

                self.calcul_data.clear()
                self.calculator_event_loop.exit()

# 58강

#### 주식 선택용 #############################################################

    def get_code_list_by_market(self, market_code):  # 종목코드들 반환
        code_list = self.dynamicCall(
            'GetCodeListByMarket(QString)', market_code)
        code_list = code_list.split(';')[:-1]
        # print(code_list) # Market 전체 code list
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
