########### Kiwoom KPI를 활용한 자동주식거래 Program ############
# kiwoom_get_code_list에서 read_code로 변경
from numpy.lib.function_base import _CORE_DIMENSION_LIST
import pandas as pd
from pandas.core.frame import DataFrame
import numpy as np
import csv
import os
import sys
import time  # roof 돌때 time.sleep(4) 4초 sleeping
# from PyQt5 import QAxContainer
# from PyQt5.QAxContainer import QAxWidget
# from PyQt5.QtTest import QTest
# from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QBoxLayout
# from PyQt5 import QtCore
# from PyQt5.QtCore import QObject

from PyQt5.QAxContainer import *  # 응용프로그램 제어용
from PyQt5.QtCore import *
from PyQt5.QtTest import *  # QTest를 통한 loop를 돌릴때 timesleep 효과

from autostock.config.errorCode import *
from autostock.config.kiwoomType import *
from autostock.config.log_class import *
# from config.slack import *


# QTest.qWait(3600) ## event loop에서 3.6초간 process를 기다려 준다.

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        self.realType = RealType()
        self.logging = Logging()
        # self.slack = Slack()  # 슬랙 동작

        print('\n===== Kiwoom Open API Start ======\n')
        self.logging.logger.debug("Kiwoom Open API logging start.")

        ##### event loop 모음 ############################
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.detail_account_info_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.calculator_event_loop = QEventLoop()
        ##################################################

        ##### 요청 스크린번호(화면번호) 모음 ############################
        self.screen_my_info = '2000'  # 계좌 관련한 스크린 번호
        self.screen_calculation_stock = '4000'  # 계산용 스크린 번호
        self.screen_real_stock = '5000'  # 종목별로 할당할 스크린 번호
        self.screen_meme_stock = '6000'  # 종목별로 할당할 주문용 스크린 번호
        self.screen_start_stop_real = '1000'  # 장 시작/종료 실시간 스크린 번호 변수

        ##################################################

        ##### 변수 모음 ############################
        # dir = 'C:/Users/histi/py37_32/autostock/' # 저장공간
        # max_day = 20 # max_day 이상 고가/저가가 이평선 위에 있는 경우. 변수 검색하여 조정
        self.account_num = None
        self.account_stock_dict = {}  # 계좌평가잔고내역요청에 따른 계좌내 종목 dict
        self.not_account_stock_dict = {}  # 미체결내역 dict
        self.portfolio_stock_dict = {}  # 관심종목 및 계좌내 종목들 dict(스크린번호 포함)
        self.jango_dict = {}  # 당일 매수한 종목들 dict

        ##################################################

        ##### 종목분석용 변수 모음 ##########################
        self.calcul_data = []

        ##### 계좌 거래 관련 변수 모음 ############################
        self.use_money = 0  # 실제 투자에 사용할 금액
        self.use_money_percent = 0.5  # 계좌에서 거래에 사용할 금액 비중.

        ##################################################

        ### 셋팅된 함수들을 순서대로 실행 ###############################
        self.get_ocx_instance()  # Kiwoom OpenKPI 접속, OCX 방식을 파이썬에 사용할 수 있게 반환
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음(이벤트 처리구역과 tr목록구역)
        self.signal_login_commConnect()  # 로그인 요청 함수 포함
        self.get_account_info()     # 계좌번호 가져오기
        self.detail_account_info()  # 예수금 요청
        self.detail_account_mystock()  # 계좌평가잔고내역 요청

        # 5초 뒤에 미체결 종목들 가져오기 실행 - 미체결종목 불러올때 오류 발생 향후 수정 요망.
        # QTimer.singleShot(5000, self.not_concluded_account)

        # my_stock_list 종목분석 실행 (임시용)
        # self.get_code_list_by_mystock('my_stock')
        # self.calculator_fnc()  # 종목분석 실행 (임시용)
        # self.read_code()  # 저장된 my_stock 불러오기.
        self.read_code('my_stock')
        self.screen_number_setting()  # 스크린 번호 할당
        self.real_event_slot()  # 실시간 거래를 위한 슬롯 연결

        QTest.qWait(2000)  # 2초 delay

        #### 주식시장 시작/완료 확인 실행 #####################################################

        # 실시간 data 요청 : FID no = 실시간목록-Real Type-장시작시간
        # 개발가이드-함수: OpenAPI.SetRealReg(_T('0150'), _T('039490'), _T('9001;302;10;11;25;12;13'), '0');  // 039490종목만 실시간 등록
        # self.dynamicCall('SetRealReg(QString, QString, QString, QString)',
        #                 self.screen_start_stop_real, '종목명', self.realType.REALTYPE['장시작시간']['장운영구분'], '0')

        self.dynamicCall('SetRealReg(QString, QString, QString, QString)',
                         self.screen_start_stop_real, '', self.realType.REALTYPE['장시작시간']['장운영구분'], '0')  # 처음 등록한 것만 '0' 이후는 모두 1로 등록

        ### 실시간에 portfolio_stock_dict에 있는 종목을 위 기초 양식에 맞춰서 등록 ##################################################

        for code in self.portfolio_stock_dict.keys():
            code_nm = self.portfolio_stock_dict[code]['종목명']
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            # 주식체결 : 1틱(주식거래체결) 발생시간.
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall(
                'SetRealReg(QString, QString, QString, QString)', screen_num, code, fids, '1')  # 추가등록이므로 1
            # print(
            #    f'실시간 관리 종목Code: {code}, 종목명:{code_nm}, 스크린번호: {screen_num}, Fids 번호: {fids}')
            self.logging.logger.debug(
                f'실시간 관리 종목Code: {code}, 종목명:{code_nm}, 스크린번호: {screen_num}, Fids 번호: {fids}')
        # print(self.portfolio_stock_dict)
        # self.slack.notification(
        #     pretext='주식자동화 프로그램 동작',
        #     title='주식 자동화 프로그램 동작',
        #     fallback='주식 자동화 프로그램 동작',
        #     text='주식 자동화 프로그램이 동작 되었습니다.'
        # )

    def get_ocx_instance(self):  # Kiwoom OpenKPI 접속
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')  # 레지스트리에 등록된 이름

    ##### Slot 목록 *개발가이드 참조 #######################################

    def event_slots(self):  # 이벤트 처리구역과 tr목록구역 생성 *개발가이드-조회와 실시간데이터 처리참조
        self.OnEventConnect.connect(self.login_slot)  # 로그인 처리 이벤트 용도.
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr(트랜잭션) 요청관련 이벤트 용도.
        # 데이터 요청 또는 주문전송 후에 서버가 보낸 메시지를 수신
        self.OnReceiveMsg.connect(self.msg_slot)

    def real_event_slot(self):  # 실시간 이벤트 구역 생성
        # 실시간 이벤트 연결  *개발가이드-조회와 실시간데이터 처리참조
        self.OnReceiveRealData.connect(self.realdata_slot)
        # 종목 주문체결 관련한 이벤트  *개발가이드-주문과 잔고처리참조
        self.OnReceiveChejanData.connect(self.chejan_slot)

    # 로그인은 CommConnect()함수를 호출하며 OnEventConnect 이벤트 인자값으로 로그인

    def signal_login_commConnect(self):
        self.dynamicCall('CommConnect()')  # 로그인 요청 시그널
        self.login_event_loop.exec_()  # 이벤트 루프 실행

    def login_slot(self, err_code):
        # print('Error 발생시 아래 error code를 확인하세요')
        # print(errors(err_code))
        self.logging.logger.debug(errors(err_code)[1])
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall(
            'GetLoginInfo(QString)', 'ACCNO')  # 계좌번호 반환
        account_num = account_list.split(';')[0]  # a;b;c   [a, b, c]

        self.account_num = account_num
        # print('\n보유계좌번호: %s' % self.account_num)  # 8158893311
        self.logging.logger.debug('보유계좌번호: %s' % self.account_num)

    def detail_account_info(self, sPrevNext='0'):  # 예수금 요청 부분

        self.dynamicCall('SetInputValue(QString, QString)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(QString, QString)', '조회구분', '2')
        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '예수금상세현황요청', 'opw00001',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

        # Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # CommRqData( 'RQName'	,  'opw00001'	,  '0'	,  '화면번호'); 위의 2000번
        # 화면번호(=스크린번호)는 grouping을 화면번호당 100개 저장 가능
        # 화면번호는 200개까지 만들수 있음.
        # 한페이지 20개 종목까지검색가능 sPrevNext =0 다음페이지없음. 2 다음페이지

    def detail_account_mystock(self, sPrevNext='0'):  # 계좌평가잔고내역요청
        # print('\n계좌평가잔고내역요청_연속조회: %s' % sPrevNext)
        self.logging.logger.debug('계좌평가잔고내역요청_연속조회: %s' % sPrevNext)

        self.dynamicCall('SetInputValue(QString, QString)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(QString, QString)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(QString, QString)', '조회구분', '2')
        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '계좌평가잔고내역요청', 'opw00018',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext='0'):  # 실시간미체결현황요청
        # print("\n실시간미체결요청")
        self.logging.logger.debug("실시간미체결요청")
        self.dynamicCall("SetInputValue(QString, QString)",
                         "계좌번호", self.account_num)
        # self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        # self.dynamicCall("SetInputValue(QString, QString)", "종목코드", "001740") #SK네트웍스
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")

        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    #### 요청에 대해서 받는 내용들 - TR 목록 #######################

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        # TR 목록의 opw0001에 있는 output 항목을 가져올 수 있음.
        if sRQName == '예수금상세현황요청':
            deposit = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '예수금')

            self.deposit = int(deposit)
            use_money = float(self.deposit) * self.use_money_percent
            self.use_money = self.use_money / 4
            self.use_money = int(use_money)

            max_order_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '주문가능금액')

            print('\n예수금: %s\n주문가능금액: %s\n사용가능금액(use_money) : %s' %
                  (format(int(deposit), ','), format(int(max_order_amount), ','), format(int(use_money), ',')))
            self.logging.logger.debug('예수금: %s 주문가능금액: %s 사용가능금액(use_money) : %s' %
                                      (format(int(deposit), ','), format(int(max_order_amount), ','), format(int(use_money), ',')))

            # self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

        # TR 목록의 opw00018에 있는 output 항목을 가져올 수 있음.
        elif sRQName == '계좌평가잔고내역요청':
            total_buy_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총매입금액')
            self.total_buy_amount = int(total_buy_amount)

            total_amount_now = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총평가금액')
            self.total_amount_now = int(total_amount_now)

            total_profit_amount = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총평가손익금액')
            self.total_profit_amount = int(total_profit_amount)

            total_profit_rate = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '총수익률(%)')
            self.total_profit_rate = float(total_profit_rate)

            print('\n총매입금액: %s\n총평가금액: %s\n총평가손익금액: %s\n총수익률(%%): %s\n' % (
                format(int(total_buy_amount), ','), format(int(total_amount_now), ','), format(int(total_profit_amount), ','), float(total_profit_rate)))
            self.logging.logger.debug('총매입금액: %s 총평가금액: %s 총평가손익금액: %s 총수익률(%%): %s' % (
                format(int(total_buy_amount), ','), format(int(total_amount_now), ','), format(int(total_profit_amount), ','), float(total_profit_rate)))

            # 한page에 20개까지만 불러오기 가능.

            rows = self.dynamicCall(
                'GetRepeatCnt(QString, QString)', sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목번호')
                code_nm = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
                learn_rate = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '수익률(%)')
                buy_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매입가')
                current_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                stock_quantity = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '보유수량')
                total_maeip_amount = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매입금액')
                current_amount = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '평가금액')
                possible_quantity = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '매매가능수량')

                code = code.strip()[1:]  # 공란을 제외하고, 두번자 글자부터 마지막까지
                code_nm = code_nm.strip()
                learn_rate = float(learn_rate.strip())
                buy_price = int(buy_price.strip())
                current_price = int(current_price.strip())
                stock_quantity = int(stock_quantity.strip())
                total_maeip_amount = int(total_maeip_amount.strip())
                current_amount = int(current_amount.strip())
                possible_quantity = int(possible_quantity.strip())

                print('종목명: %s\n수익률(%%): %s\n매입가: %s\n현재가: %s\n보유수량: %s\n매입금액: %s\n평가금액: %s\n' % (
                    code_nm, learn_rate, format(buy_price, ','), format(current_price, ','), format(stock_quantity, ','), format(total_maeip_amount, ','), format(current_amount, ',')))
                self.logging.logger.debug('종목명: %s 수익률(%%): %s 매입가: %s 현재가: %s 보유수량: %s 매입금액: %s 평가금액: %s' % (
                    code_nm, learn_rate, format(buy_price, ','), format(current_price, ','), format(stock_quantity, ','), format(total_maeip_amount, ','), format(current_amount, ',')))

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
                asd.update({'매매가능수량': possible_quantity})

            print('내 계좌에 있는 종목수: %s\n' % len(self.account_stock_dict))
            self.logging.logger.debug(
                '내 계좌에 있는 종목수: %s' % len(self.account_stock_dict))

            if sPrevNext == '2':
                self.detail_account_mystock(sPrevNext='2')
            else:
                self.detail_account_info_event_loop.exit()

        # TR 목록의 opt10075에 있는 output 항목을 가져올 수 있음.
        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall(
                "GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            print('미체결에 있는 종목수: %s\n' % rows)
            self.logging.logger.debug('미체결에 있는 종목수: %s' % rows)

            for i in range(rows):
                code = self.dynamicCall(
                    'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목코드')
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

                print(f'sRQName == 실시간미체결요청 code 확인용:{code}')
                self.logging.logger.debug(
                    f'sRQName == 실시간미체결요청 code 확인용:{code}')

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

                print('미체결종목: %s\n' % self.not_account_stock_dict[code_nm])
                self.logging.logger.debug(
                    '미체결종목: %s\n' % self.not_account_stock_dict[code_nm])

            self.detail_account_info_event_loop.exit()

        # TR 목록의 opt10081에 있는 output 항목을 가져올 수 있음.

        elif sRQName == '주식일봉차트조회':

            code = self.dynamicCall(
                'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '종목코드')
            # print(code)
            code = code.strip()
            # data = self.dynamicCall('GetCommDataEx(QString, QString)', sTrCode, sRQName)

            print('%s 주식일봉데이터요청' % code)
            self.logging.logger.debug('%s 주식일봉데이터요청' % code)

            cnt = self.dynamicCall(
                'GetRepeatCnt(QString, QString)', sTrCode, sRQName)

            print('남은 데이터 일자 수 %s' % cnt)
            self.logging.logger.debug('남은 데이터 일자 수 %s' % cnt)

            # 한번조회할때 600일치까지 일봉 data를 받을 수 있다.
            daily_df = pd.DataFrame(
                [], columns=['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가', ''])
            for i in range(cnt):
                # data = self.dynamicCall('GetCommData(QString, QString)', sTrCode, sRQName)
                # 결과물 [['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가',''],['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가','']]
                data = []

                current_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
                value = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '거래량')
                trading_value = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '거래대금')
                date = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '일자')
                start_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '시가')
                high_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '고가')
                low_price = self.dynamicCall(
                    'GetCommData(QString, QString, int, QString)', sTrCode, sRQName, i, '저가')

                data.append('')
                data.append(current_price.strip())
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append('')

                self.calcul_data.append(data.copy())
                # calcul_data 내용 [['', '현재가', '거래량', '거래금액', '일자', '시가', '고가', '저가',''], [....]]

                # daily_df = daily_df.append(
                #     daily_df.iloc[-1], ignore_index=True)
                # daily_df.iloc[-1] = self.calcul_data
                # print(f'daily_df 길이: {len(daily_df)}')

            if sPrevNext == '2':
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)

            else:
                print('총 일수 %s' % len(self.calcul_data))
                self.logging.logger.debug('총 일수 %s' % len(self.calcul_data))

############### 종목 선정 조건 설정 #################################
############### 1. 그랜빌의 매수신호 4법칙 계산 ######################

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
                        self.logging.logger.debug('오늘 주가 120일 이평선에 걸쳐 있는지 확인')
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
                                self.logging.logger.debug('120일치 data 없음')
                                break

                            total_price = 0
                            for value in self.calcul_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_average_price_prev = total_price / 120

                            # max_day이상 고가가 120일 이평선 위에 있을 경우 제외
                            max_day = 60
                            if int(self.calcul_data[idx][6]) >= moving_average_price_prev and idx <= max_day:
                                print('### 고가가 %s일 이상 120일 이평선 위에 있어서 제외됨' %
                                      max_day)
                                self.logging.logger.debug('### 고가가 %s일 이상 120일 이평선 위에 있어서 제외됨' %
                                                          max_day)
                                price_top_moving = False
                                break

                            # 저가가 120일 이평선 위에 있었다면 제외
                            elif int(self.calcul_data[idx][7]) > moving_average_price_prev and idx > max_day:
                                print('### 저가가 %s일 이상 120일 이평선 위에 있어서 제외됨' %
                                      max_day)
                                self.logging.logger.debug(
                                    '### 저가가 %s일 이상 120일 이평선 위에 있어서 제외됨' % max_day)
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
                                self.logging.logger.debug(
                                    '### 포착된 이평선 가격이 오늘자 이평선 가격보다 낮고')
                                self.logging.logger.debug(
                                    '### 과거 일봉 저가가 오늘자 일봉의 고가보다 낮은 것 확인됨')
                                pass_success = True

                if pass_success == True:
                    print('조건부 통과')
                    self.logging.logger.debug('조건부 통과')

                    code_nm = self.dynamicCall(
                        'GetMasterCodeName(QString)', code)

                    # a는 연결해서 쓴다. w는 덮어쓴다.
                    dir = 'C:/Users/histi/py37_32/autostock/'
                    f = open(f'{dir}files/그랜빌_stock.txt',
                             'a', encoding='utf8')
                    f.write(
                        f'{code}\t{code_nm}\t{str(self.calcul_data[0][1])}\n')
                    # f.write('%s\t%s\t%s\n' %
                    #         (code, code_nm, str(self.calcul_data[0][1])))
                    f.close()

                elif pass_success == False:
                    print('조건부 통과 못함\n')
                    self.logging.logger.debug('조건부 통과 못함')

                self.calcul_data.clear()
                self.calculator_event_loop.exit()

#### 주식 선택용(강의) #############################################################

    # def get_code_list_by_market(self, market_code):  # 종목코드들 반환
    #     code_list = self.dynamicCall(
    #         'GetCodeListByMarket(QString)', market_code)
    #     code_list = code_list.split(';')[:-1]
    #     # print(code_list) # Market 전체 code list
    #     return code_list

    # def calculator_fnc(self):  # 종목분석 실행용 함수
    #     # code_list = self.get_code_list_by_market('0') #코스피 마켓
    #     code_list = self.get_code_list_by_market('10')  # 코스닥 마켓
    #     print('코스닥 종목 갯수: %s' % len(code_list))
    #     #self.logging.logger.debug('코스닥 종목 갯수: %s ' % len(code_list))

    #     # code_list 안에서 index와 value를 모두 사용.enumerate
    #     for idx, code in enumerate(code_list):
    #         self.dynamicCall('DisconnectRealData(QString)',
    #                          self.screen_calculation_stock)  # 스크린 연결 끊기
    #         print('%s / %s : KOSDAQ Stock Code : %s is updating... ' %
    #               (idx + 1, len(code_list), code))
    #         #self.logging.logger.debug('%s / %s : KOSDAQ Stock Code : %s is updating... ' % (idx + 1, len(code_list), code))

    #         self.day_kiwoom_db(code=code)

#### 주식 선택용(My stock list사용) #####################################################

    # ### 하나의 관심섹터의 종목만 ####
    # def get_code_list_by_mystock(self, my_stock):  # 나의 관심종목코드를 불러와서 사용
    #     # my_stock = 'my_stock'
    #     dir = 'C:/Users/histi/py37_32/autostock/'
    #     df = pd.read_excel(f'{dir}files/{my_stock}.xls', dtype={'종목코드': str},) # sheet_name=0 특정sheet만
    #     df = df.drop(['Unnamed: 0', '매입단가', '매입수량', '메모'],
    #                  axis=1).dropna(how='all', axis=0)
    #     return df

    ### 모든 관심섹터의 종목 ####
    # def get_code_list_by_mystock(self, my_stock):  # 나의 관심종목코드를 불러와서 사용
    #     my_stock = 'my_stock'
    #     dir = 'C:/Users/histi/py37_32/autostock/'
    #     df_all = pd.read_excel(
    #         f'{dir}files/{my_stock}.xls', sheet_name=None, dtype={'종목코드': str})
    #     df_all.keys()
    #     concatted_df = pd.concat(df_all)
    #     concatted_df.columns
    #     df = concatted_df.drop(
    #         ['Unnamed: 0', '매입단가', '매입수량', '메모'], axis=1).dropna(how='all', axis=0)
    #     df = df.drop_duplicates(subset=['종목코드'], keep='first')
    #     df = df.dropna(how='all', axis=1).reset_index()
    #     return df

    def calculator_fnc(self):  # 종목분석 실행용 함수
        # code_list = self.get_code_list_by_mystock('my_stock')
        code_list = self.read_code('my_stock')
        print('나의 관심 종목 갯수: %s종목' % len(code_list))
        print(f'나의 관심 종목 list\n{code_list}')
        #self.logging.logger.debug('나의 관심 종목 갯수: %s' % len(code_list))

        for idx in code_list.index:
            code = code_list.loc[idx, '종목코드']
            print(
                f'{idx+1} / {len(code_list)} : my Stock list Code : {code} is updating... ')
            # idx += 1
            self.logging.logger.debug(
                f'{idx+1} / {len(code_list)} : my Stock list Code : {code} is updating... ')
            self.day_kiwoom_db(code=code)

########################################################################################

    def day_kiwoom_db(self, code=None, date=None, sPrevNext='0'):  # 주식일봉차트조회
        QTest.qWait(3600)  # 3.6초마다 딜레이를 준다..

        self.dynamicCall('SetInputValue(QString, QString)', '종목코드', code)
        self.dynamicCall('SetInputValue(QString, QString)', '수정주가구분', '1')

        if date != None:
            self.dynamicCall('SetInputValue(QString, QString)', '기준일자', date)

        self.dynamicCall('CommRqData(QString, QString, int, QString)',
                         '주식일봉차트조회', 'opt10081', sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec_()

#### My stock list 나의 관심 종목 불러오기 ############################################
    # def read_code(self):
    def read_code(self, my_stock):
        my_stock = 'my_stock'
        dir = 'C:/Users/histi/py37_32/autostock/'
        df_all = pd.read_excel(
            f'{dir}files/{my_stock}.xls', sheet_name=None, dtype={'종목코드': str})
        df_all.keys()
        concatted_df = pd.concat(df_all)
        concatted_df.columns
        df = concatted_df.drop(
            ['Unnamed: 0', '매입단가', '매입수량', '메모'], axis=1).dropna(how='all', axis=0)
        df = df.drop_duplicates(subset=['종목코드'], keep='first')
        df = df.dropna(how='all', axis=1).reset_index()
        stock_code = df['종목코드'].values
        stock_n = df['종목명'].values

        stock_name = []
        for i in stock_n:
            stock_name.append({'종목명': i})
        self.portfolio_stock_dict = dict(zip(stock_code, stock_name))
        return df

    def screen_number_setting(self):
        screen_overwite = []
        # 계좌평가잔고 내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwite:
                screen_overwite.append(code)
        # 미체결에 있는 종목들
        for order_number in self.not_account_stock_dict.keys():
            code = self.not_account_stock_dict[order_number]['종목코드']
            if code not in screen_overwite:
                screen_overwite.append(code)
        # 포트폴리오에 담겨있는 종목들
        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwite:
                screen_overwite.append(code)

        # 스크린 번호 할당
        cnt = 0
        for code in screen_overwite:

            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)
            # print(code, temp_screen, meme_screen)

            if (cnt % 50) == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)

            if (cnt % 50) == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)

            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update(
                    {'스크린번호': str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update(
                    {'주문용스크린번호': self.screen_meme_stock})

            elif code not in self.portfolio_stock_dict.keys():  # 종목명 없음.
                # self.portfolio_stock_dict[code].update(
                #     {'종목명': self.account_stock_dict.keys})
                self.portfolio_stock_dict.update(
                    {code: {'스크린번호': str(self.screen_real_stock), '주문용스크린번호': str(self.screen_meme_stock)}})
                # self.portfolio_stock_dict.update(
                #     {code: {'주문용스크린번호': str(self.screen_meme_stock)}})

            cnt += 1

        # print(self.portfolio_stock_dict)
# my stock list Dict 저장 code:{'225190': {'종목명': '삼양옵틱스', '스크린번호': '5001', '주문용스크린번호': '6001'}

        dir = 'C:/Users/histi/py37_32/autostock/'
        f = open(f'{dir}files/portforlio_stock.txt',
                 'w', encoding='utf8')
        f.write(
            f'{str(self.portfolio_stock_dict)}')
        f.close()

#### 실시간 data 불러오기 ##################################################

#### 1. 장 시작 , 운영구분 등록 불러오기 ####################################

    def realdata_slot(self, sCode, sRealType, sRealData):
        # (0:장시작전, 2:장종료전(20분), 3:장시작, 4,8:장종료(30분), 9:장마감)
        if sRealType == '장시작시간':
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            value = self.dynamicCall(
                'GetCommRealData(QString, int)', sCode, fid)
            print(f'장운영구분_no:{fid}, fid_no={value}')
            self.logging.logger.debug(f'장운영구분_no:{fid}, fid_no={value}')

            if value == '0':
                print('장 시작 전')
                self.logging.logger.debug('장 시작 전')

            elif value == '3':
                print('장 시작')
                self.logging.logger.debug('장 시작')

            elif value == '2':
                print('장 종료, 동시호가로 넘어감')
                self.logging.logger.debug('장 종료, 동시호가로 넘어감')

# 전체 계좌 종목 일괄매도 프로그램 계발 예정 ===========================================

            elif value == '4':
                print('3시30분 장 종료')
                self.logging.logger.debug('3시30분 장 종료')

                # 장마감 후, 사용한 list 모두 삭제
                for code in self.portfolio_stock_dict.keys():
                    self.dynamicCall("SetRealRemove(QString, QString)",
                                     self.portfolio_stock_dict[code]['스크린번호'], code)

                # 장마감 후, 사용한 list 모두 삭제 slot 실행
                # self.file_delete()

                # 장마감 후, 종목선정을 위한 calculator slot 실행
                # self.calculator_fnc()

                # 파이썬에서 사용된 모든 library 삭제.
                sys.exit()

#### 2. 관심 종목들 주식체결 실시간 틱 data 불러와서 업데이트 ################################

        elif sRealType == '주식체결':
            QTest.qWait(100)  # 0.1초마다 딜레이를 준다..
            a = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['체결시간'])  # 출력 HHMMSS
            b = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['현재가'])  # 출력 : +(-)2520
            b = abs(int(b))

            c = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['전일대비'])  # 출력 : +(-)2520
            c = abs(int(c))

            d = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['등락율'])  # 출력 : +(-)12.98
            d = float(d)

            e = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['(최우선)매도호가'])  # 출력 : +(-)2520
            e = abs(int(e))

            f = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['(최우선)매수호가'])  # 출력 : +(-)2515
            f = abs(int(f))

            g = self.dynamicCall('GetCommRealData(QString, int)',
                                 sCode, self.realType.REALTYPE[sRealType]['거래량'])  # 틱 하나당 거래량
            g = abs(int(g))

            h = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['누적거래량'])  # 출력 : 240124
            h = abs(int(h))

            i = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['고가'])  # 출력 : +(-)2530
            i = abs(int(i))

            j = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['시가'])  # 출력 : +(-)2530
            j = abs(int(j))

            k = self.dynamicCall('GetCommRealData(QString, int)', sCode,
                                 self.realType.REALTYPE[sRealType]['저가'])  # 출력 : +(-)2530
            k = abs(int(k))

            self.portfolio_stock_dict[sCode].update({'체결시간': a})
            self.portfolio_stock_dict[sCode].update({'현재가': b})
            self.portfolio_stock_dict[sCode].update({'전일대비': c})
            self.portfolio_stock_dict[sCode].update({'등락율': d})
            self.portfolio_stock_dict[sCode].update({'(최우선)매도호가': e})
            self.portfolio_stock_dict[sCode].update({'(최우선)매수호가': f})
            self.portfolio_stock_dict[sCode].update({'거래량': g})
            self.portfolio_stock_dict[sCode].update({'누적거래량': h})
            self.portfolio_stock_dict[sCode].update({'고가': i})
            self.portfolio_stock_dict[sCode].update({'시가': j})
            self.portfolio_stock_dict[sCode].update({'저가': k})

############## code 포함한 log 만들기 ###################################

            # print(self.portfolio_stock_dict[sCode]) #key값을 제외하고 출력
            # self.logging.logger.debug(self.portfolio_stock_dict[sCode]) #key값을 제외하고 출력

            print(self.portfolio_stock_dict)
            # self.logging.logger.debug(self.portfolio_stock_dict)

            dir = 'C:/Users/histi/py37_32/autostock/files/tic_data/'
            day = format(datetime.now(), '%Y-%m-%d')
            # t_d = {sCode : {'date' : day})
            #portfolio_stock_tic_data = self.portfolio_stock_dict[sCode].update({'date': t_d})

            f = open(f'{dir}{day}_v1.txt', 'a', encoding='utf8')
            f1 = open(f'{dir}{day}_v2.txt', 'a', encoding='utf8')

            f.write(f'{str(self.portfolio_stock_dict)}\n')
            f1.write(f'{str(self.portfolio_stock_dict)}')
            # f.write(f'{str(portfolio_stock_tic_data)}\n')
        # f.close()
        # f1.close()

#### 2. 주식 주문을 위한 조건문 구성 ####################################
        # self.portfolio_stock_dict[sCode].update({'체결시간': a})
        # self.portfolio_stock_dict[sCode].update({'현재가': b})
        # self.portfolio_stock_dict[sCode].update({'전일대비': c})
        # self.portfolio_stock_dict[sCode].update({'등락율': d})
        # self.portfolio_stock_dict[sCode].update({'(최우선)매도호가': e})
        # self.portfolio_stock_dict[sCode].update({'(최우선)매수호가': f})
        # self.portfolio_stock_dict[sCode].update({'거래량': g})
        # self.portfolio_stock_dict[sCode].update({'누적거래량': h})
        # self.portfolio_stock_dict[sCode].update({'고가': i})
        # self.portfolio_stock_dict[sCode].update({'시가': j})
        # self.portfolio_stock_dict[sCode].update({'저가': k})

# 매수, 매도 가격조건 설정 =====================================================

        buy_price = 0  # 매도주문 가격 조건 선택:시장가
        order_price = 0  # 매수주문 가격 조건 선택: 현재가 = b
        # order_price = f # 매수주문 가격 조건 선택: '(최우선)매수호가'


# 매수, 손절, 익절 % 설정 =====================================================

        # buying_up_rate = 1.0  # 신규 매수 기준 등락율이 1% 이상 종목, 현재 잔고에 없을 경우.
        buying_down_rate = -4.5  # 신규 매수 기준 등락율이 -4.5% 이상 종목, 현재 잔고에 없을 경우.
        expected_profit_rate = 4.0  # 수익실현
        expected_loss_rate = -2.0  # 손절

# 계좌잔고평가내역(account_stock_dict)에 있고 오늘 매수한 잔고에 없는 경우 ==========
        if sCode in self.account_stock_dict.keys() and sCode not in self.jango_dict.keys():
            asd = self.account_stock_dict[sCode]
            code_nm = asd['종목명']
            print(self.account_stock_dict)
            self.logging.logger.debug(self.account_stock_dict)

            meme_rate = (b - asd['매입가']) / asd['매입가'] * 100  # 등락율

            # LONG nPrice -> 0 이면 시장가로 주문.
            if asd['매매가능수량'] > 0 and (meme_rate > expected_profit_rate or meme_rate < expected_loss_rate):
                order_success = self.dynamicCall(
                    'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                    ['신규매도', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,
                        2, sCode, asd['매매가능수량'], buy_price, self.realType.SENDTYPE['거래구분']['시장가'], ""]
                )

                if order_success == 0:
                    print(f'계좌잔고 종목: {sCode}, {code_nm}, 매도 주문 성공')
                    print('주문시점 meme_rate = %s' %
                          format(round(meme_rate, 2)))
                    self.logging.logger.debug(
                        f'계좌잔고 종목: {sCode}, {code_nm}, 매도 주문 성공')
                    self.logging.logger.debug('주문시점 meme_rate = %s' %
                                              format(round(meme_rate, 2)))
                    del self.account_stock_dict[sCode]
                else:
                    print(f'계좌잔고 종목: {sCode}, {code_nm}, 매도 주문 실패')
                    self.logging.logger.debug(
                        f'계좌잔고 종목: {sCode}, {code_nm}, 매도 주문 실패')

                QTest.qWait(100)  # 0.1초

# 오늘 매입한 종목. 잔고에 있을 경우 =====================================================

        elif sCode in self.jango_dict.keys():
            jd = self.jango_dict[sCode]
            code_nm = jd['종목명']
            print(self.jango_dict)
            self.logging.logger.debug(self.jango_dict)

            meme_rate = (b - jd['매입단가']) / jd['매입단가'] * 100  # 등락율

            if jd['주문가능수량'] > 0 and (meme_rate > expected_profit_rate or meme_rate < expected_loss_rate):
                order_success = self.dynamicCall(
                    'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                    ['신규매도', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,
                        2, sCode, jd['주문가능수량'], buy_price, self.realType.SENDTYPE['거래구분']['시장가'], '']
                )

                if order_success == 0:
                    print(f'오늘매수한 종목: {sCode}, {code_nm}, 매도 주문 성공')
                    print('주문시점 meme_rate = %s' %
                          format(round(meme_rate, 2)))
                    self.logging.logger.debug(
                        f'오늘매수한 종목: {sCode}, {code_nm}, 매도 주문 성공')
                    self.logging.logger.debug('주문시점 meme_rate = %s' %
                                              format(round(meme_rate, 2)))
                else:
                    print(f'오늘매수한 종목: {sCode}, {code_nm}, 매도 주문 실패')
                    self.logging.logger.debug(
                        f'오늘매수한 종목: {sCode}, {code_nm}, 매도 주문 실패')

                QTest.qWait(100)  # 0.1초

# 매수 실행 : d = '등락율', b = '현재가' =====================================================
# 당일 신규 매수 기준 등락율이 buying_up_rate % 제외
# 당일 관심종목 buying_down_rate % 미만, 현재 계좌와 당일 매수 잔고에 있을 경우 매수제외

            elif d < buying_down_rate and sCode in self.account_stock_dict:
                # self.logging.logger.debug('매수조건 통과 %s ' % sCode)
                code_nm = self.portfolio_stock_dict[sCode]['종목명']
                # print(
                #    f'self.account_stock_dict 내용은 : {self.account_stock_dict}')
                print(f'{code_nm} 매수조건 충족했지만 계좌에 있는 종목, 매수 제한입니다.')
                self.logging.logger.debug(
                    f'{code_nm} 매수조건 충족했지만 계좌에 있는 종목, 매수 제한입니다.')

                QTest.qWait(100)  # 0.1초

            elif d < buying_down_rate and sCode in self.jango_dict:
                code_nm = self.portfolio_stock_dict[sCode]['종목명']
                # print(
                #    f'self.jango_dict 내용은 : {self.jango_dict}')
                print(f'{code_nm} 매수조건 충족했지만 당일 매수잔고에 있는 종목, 매수 제한입니다.')
                self.logging.logger.debug(
                    f'{code_nm} 매수조건 충족했지만 당일 매수잔고에 있는 종목, 매수 제한입니다.')

                QTest.qWait(100)  # 0.1초

# 신규 매수 기준 등락율이 buying_up_rate % 이상 or buying_down_rate 미만인 종목, 현재 잔고에 없을 경우.
# 신규 매수 포함 종목당 총잔고가 500만원 초과할 경우, 매수 한계 표시.
                # d = '등락율' b = '현재가'
            elif d < buying_down_rate and sCode not in self.jango_dict:
                code_nm = self.portfolio_stock_dict[sCode]['종목명']

                # print(f'self.jango_dict 내용은 : {self.jango_dict}')

                # order_price = 112000
                # use_money = 2467861
                # code_nm = 'TEST'

                result = abs(int(self.use_money * 0.01))
                quantity = int((result) / order_price)

                print('신규매수 종목: %s\n매수주문 금액: %s\n매수주문 수량 : %s\n현재가: %s' %
                      (code_nm, format((result), ','), format((quantity), ','), format((order_price), ',')))
                self.logging.logger.debug('신규매수 종목: %s 매수주문 금액: %s 매수주문 수량: %s 현재가: %s' %
                                          (code_nm, format((result), ','), format((quantity), ','), format((order_price), ',')))

                order_success = self.dynamicCall(
                    "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                    ["신규매수", self.portfolio_stock_dict[sCode]["주문용스크린번호"], self.account_num,
                        1, sCode, quantity, order_price, self.realType.SENDTYPE['거래구분']['지정가'], ""]
                )

                if order_success == 0:
                    print(
                        f'신규매수 종목:{sCode}, {code_nm}, {quantity}주 매수 주문 성공')
                    print('신규매수 주문시점 등락율 = %s' % (format(round(d, 2))))
                    self.logging.logger.debug(
                        f'신규매수 종목:{sCode}, {code_nm}, {quantity}주 매수 주문 성공')
                    self.logging.logger.debug(
                        '신규매수 주문시점 등락율 = %s' % (format(round(d, 2))))

                else:
                    print(
                        f'신규매수 종목:{sCode}, {code_nm}, {quantity}주 매수 주문 실패')
                    self.logging.logger.debug(
                        f'신규매수 종목:{sCode}, {code_nm}, {quantity}주 매수 주문 실패')

                QTest.qWait(100)  # 0.1초

# 매매 list 작성, meme_list update  =====================================================================

            not_meme_list = list(self.not_account_stock_dict)
            # not_meme_list = self.not_account_stock_dict.copy() 상동.data를 새로운 주소에 copy해서 사용한다.
            # 아래 처럼하면 두개의 list의 data가 같은 주소에 배치되므로 not_account_stock_dict data가 실시간 변동되는 경우, 조건문등에서 error가 발생한다.
            # not_meme_list = self.not_account_stock_dict

            for order_num in not_meme_list:
                code = self.not_account_stock_dict[order_num]['종목코드']
                code_nm = self.not_account_stock_dict[order_num]['종목명']
                meme_price = self.not_account_stock_dict[order_num]['주문가격']
                not_quantity = self.not_account_stock_dict[order_num]['미체결수량']
                order_gubun = self.not_account_stock_dict[order_num]['주문구분']

                if order_gubun == '매수' and not_quantity > 0 and buy_price > meme_price:  # 아래 0은 전체 취소.
                    print('매수취소한다')
                    self.logging.logger.debug('매수취소한다')
                    order_success = self.dynamicCall(
                        'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                        ['매수취소', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,
                            3, code, 0, 0, self.realType.SENDTYPE['거래구분']['지정가'], order_num]
                    )

                    if order_success == 0:
                        print(f'{code}, {code_nm} 매수취소 전달 성공')
                        self.logging.logger.debug(
                            f'{code}, {code_nm} 매수취소 전달 성공')
                    else:
                        print(f'{code}, {code_nm}-매수취소 전달 실패')
                        self.logging.logger.debug(
                            f'{code}, {code_nm}-매수취소 전달 실패')

                elif not_quantity == 0:   # 미체결 수량이 0 면 삭제. 단타의 경우, for문을 위해서 삭제.
                    del self.not_account_stock_dict[order_num]

# 개발가이드 주문 잔고-함수-void OnRreceiveChejanData() 참조. ==========================

    def chejan_slot(self, sGubun, nItemCnt, sFidList):
        # 실시간목록 - 주문체결, 잔고 항목에서 가져옴.
        # sGubun = 주문체결 0, 잔고변경 1 -> 실시간 목록 주문체결, 잔고에서 아래 사항 FID 확인가능.
        if int(sGubun) == 0:
            account_num = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['계좌번호'])
            sCode = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['종목코드'])[1:]  # 체결의 종목코드 첫글자에 영문자가 포함되어져 있으므로 제외처림.
            stock_name = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['종목명'])
            stock_name = stock_name.strip()

            origin_order_number = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['원주문번호'])  # 출력 : defaluse : '000000'
            order_number = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문번호'])  # 출럭: 0115061 마지막 주문번호

            order_status = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문상태'])  # 출력: 접수, 확인, 체결
            order_quan = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문수량'])  # 출력 : 3
            order_quan = int(order_quan)

            order_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문가격'])  # 출력: 21000
            order_price = int(order_price)

            not_chegual_quan = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['미체결수량'])  # 출력: 15, default: 0
            not_chegual_quan = int(not_chegual_quan)

            order_gubun = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문구분'])  # 출력: -매도, +매수
            # 주문구분에 + - 삭제.
            order_gubun = order_gubun.strip().lstrip('+').lstrip('-')

            chegual_time_str = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문/체결시간'])  # 출력: '151028'

            chegual_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['체결가'])  # 출력: 2110 default : ''
            if chegual_price == '':
                chegual_price = 0
            else:
                chegual_price = int(chegual_price)

            chegual_quantity = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['체결량'])  # 출력: 5 default : ''
            if chegual_quantity == '':
                chegual_quantity = 0
            else:
                chegual_quantity = int(chegual_quantity)

            current_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['현재가'])  # 출력: -6000
            current_price = abs(int(current_price))

            first_sell_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['(최우선)매도호가'])  # 출력: -6010
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['(최우선)매수호가'])  # 출력: -6000
            first_buy_price = abs(int(first_buy_price))

            # 새로 들어온 주문이면 dict에 주문번호 할당
            if order_number not in self.not_account_stock_dict.keys():
                self.not_account_stock_dict.update({order_number: {}})

            self.not_account_stock_dict[order_number].update({'종목코드': sCode})
            self.not_account_stock_dict[order_number].update(
                {'주문번호': order_number})
            self.not_account_stock_dict[order_number].update(
                {'종목명': stock_name})
            self.not_account_stock_dict[order_number].update(
                {'주문상태': order_status})
            self.not_account_stock_dict[order_number].update(
                {'주문수량': order_quan})
            self.not_account_stock_dict[order_number].update(
                {'주문가격': order_price})
            self.not_account_stock_dict[order_number].update(
                {'미체결수량': not_chegual_quan})
            self.not_account_stock_dict[order_number].update(
                {'원주문번호': origin_order_number})
            self.not_account_stock_dict[order_number].update(
                {'주문구분': order_gubun})
            self.not_account_stock_dict[order_number].update(
                {'주문/체결시간': chegual_time_str})
            self.not_account_stock_dict[order_number].update(
                {'체결가': chegual_price})
            self.not_account_stock_dict[order_number].update(
                {'체결량': chegual_quantity})
            self.not_account_stock_dict[order_number].update(
                {'현재가': current_price})
            self.not_account_stock_dict[order_number].update(
                {'(최우선)매도호가': first_sell_price})
            self.not_account_stock_dict[order_number].update(
                {'(최우선)매수호가': first_buy_price})

            # print(self.not_account_stock_dict)  # 수동 입력 결과 확인 할것.

        elif int(sGubun) == 1:  # 잔고
            account_num = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['계좌번호'])
            sCode = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['종목코드'])[1:]

            stock_name = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['종목명'])
            stock_name = stock_name.strip()

            current_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['현재가'])
            current_price = abs(int(current_price))

            stock_quan = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['보유수량'])
            stock_quan = int(stock_quan)

            like_quan = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['주문가능수량'])
            like_quan = int(like_quan)

            buy_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['매입단가'])  # 평균매입단가
            buy_price = abs(int(buy_price))

            total_buy_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['총매입가'])  # 계���������에 있는 종목의 총매입가
            total_buy_price = int(total_buy_price)

            meme_gubun = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['매도매수구분'])
            meme_gubun = self.realType.REALTYPE['매도수구분'][meme_gubun]

            first_sell_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['잔고']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            if sCode not in self.jango_dict.keys():
                self.jango_dict.update({sCode: {}})

            self.jango_dict[sCode].update({'현재가': current_price})
            self.jango_dict[sCode].update({'종목코드': sCode})
            self.jango_dict[sCode].update({'종목명': stock_name})
            self.jango_dict[sCode].update({'보유수량': stock_quan})
            self.jango_dict[sCode].update({'주문가능수량': like_quan})
            self.jango_dict[sCode].update({'매입단가': buy_price})
            self.jango_dict[sCode].update({'총매입가': total_buy_price})
            self.jango_dict[sCode].update({'매도매수구분': meme_gubun})
            self.jango_dict[sCode].update({'(최우선)매도호가': first_sell_price})
            self.jango_dict[sCode].update({'(최우선)매수호가': first_buy_price})

            if stock_quan == 0:  # 미체결계좌에서 사라지므로 dict에서 제외, 실시간 종목 연결 제외
                del self.jango_dict[sCode]
                self.dynamicCall('SetRealRemove(QString, QString)',
                                 self.portfolio_stock_dict[sCode]['스크린번호'], sCode)

# 송수신 메세지 get slot 생성
    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        print('스크린: %s, 요청이름: %s, tr코드: %s, ----- %s' %
              (sScrNo, sRQName, sTrCode, msg))
        self.logging.logger.debug('스크린: %s, 요청이름: %s, tr코드: %s, ----- %s' %
                                  (sScrNo, sRQName, sTrCode, msg))


# 장마감 후, 생성 화일 삭제
    # def file_delete(self):
    #     if os.path.isfile('f'{dir}files/portforlio_stock.txt'):
    #         os.remove(f'{dir}files/portforlio_stock.txt')


# Finish ========================================================================

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


# 개발가이드-주문과 잔고처리-관련함수-LONG SendOrder 참조.
#   SendOrder(
#   BSTR sRQName, // 사용자 구분명
#   BSTR sScreenNo, // 화면번호
#   BSTR sAccNo,  // 계좌번호 10자리
#   LONG nOrderType,  // 주문유형 1:신규매수, 2:신규매도 3:매수취소, 4:매도취소, 5:매수정정, 6:매도정정
#   BSTR sCode, // 종목코드 (6자리)
#   LONG nQty,  // 주문수량
#   LONG nPrice, // 주문가격
#   BSTR sHogaGb,   // 거래구분(혹은 호가구분)은 아래 참고
#   BSTR sOrgOrderNo  // 원주문번호. 신규주문에는 공백 입력, 정정/취소시 입력합니다.
#   )


# 종목명 : LG디스플레이
# 신규매수주문 금액 : 9,963,776
# 신규매수주문 수량 : 7,380
# 034220-LG디스플레이-7380주 매수 주문 전달 성공
# 계좌내 LG디스플레이, buying_down_rate = -4.5
# 스크린: 6001, 요청이름: 신규매수, tr코드: KOA_NORMAL_BUY_KP_ORD, ----- [00Z218] 모의투
# 자 장종료 상태입니다
