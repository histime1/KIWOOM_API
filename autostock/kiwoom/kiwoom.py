from PyQt5.QAxContainer import *  # 응용프로그램 제어용
from PyQt5.QtCore import *
from autostock.config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom Class입니다.')
        ##### event loop 모음 ############################
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        ##################################################

        ##### 스크린번호(화면번호) 모음 ############################
        self.screen_my_info = '2000'

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

        ### 실행할 모든 함수명 모음 ###############################
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()  # 예수금요청 부분
        self.detail_account_mystock()  # 계좌평가잔고내역요청
        self.not_concluded_account()  # 실시간매체결현황요청
        ##################################################

    def get_ocx_instance(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')  # 레지스트리에 등록된 이름

    def event_slots(self):  # 이벤트 처리구역
        self.OnEventConnect.connect(self.login_slot)  # 로그인 처리 진행용.
        self.OnReceiveTrData.connect(self.trdata_slot)  # tr 목록 data 가져오는 용도

    # 로그인은 CommConnect()함수를 호출하며 OnEventConnect 이벤트 인자값으로 로그인

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def login_slot(self, errCode):
        print(errors(errCode))

        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall('GetLogininfo(String)', 'ACCNO')
        self.account_num = account_list.split(';')[0]
        print('나의 보유계좌번호: %s' % self.account_num)  # 8158893311

    def detail_account_info(self):
        print('예수금 요청 부분')

        self.dynamicCall('SetInputValue(String, String)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(String, String)', '조회구분', '2')
        self.dynamicCall('CommRqData(String, String, int, String)',
                         '예수금상세현황요청', 'opw00001',  '0',  self.screen_my_info)
        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()
        # Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # CommRqData( "RQName"	,  "opw00001"	,  "0"	,  "화면번호"); 위의 2000번
        # 화면번호(=스크린번호)는 grouping을 화면번호당 100개 저장 가능
        # 화면번호는 200개까지 만들수 있음.

    # 한페이지 20개 종목까지검색가능 sPrevNext =0 다음페이지없음. 2 다음페이지
    def detail_account_mystock(self, sPrevNext='0'):
        print('계좌평가잔고내역요청_연속조회: %s' % sPrevNext)

        self.dynamicCall('SetInputValue(String, String)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '비밀번호', '0269')
        self.dynamicCall('SetInputValue(String, String)', '비밀번호입력매체구분', '00')
        self.dynamicCall('SetInputValue(String, String)', '조회구분', '2')
        self.dynamicCall('CommRqData(String, String, int, String)',
                         '계좌평가잔고내역요청', 'opw00018',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext='0'):
        print('실시간미체결현황요청')
        self.dynamicCall('SetInputValue(String, String)',
                         '계좌번호', self.account_num)
        self.dynamicCall('SetInputValue(String, String)', '매매구분', '0')
        self.dynamicCall('SetInputValue(String, String)', '체결구분', '1')
        self.dynamicCall('CommRqData(String, String, int, String)',
                         '실시간미체결현황', 'opt10075',  sPrevNext,  self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

########################## 요청에 대해서 받는 내용들 - TR 목록 #######################

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == '예수금상세현황요청':  # TR 목록의 opw0001에 있는 output 항목을 가져올 수 있음.
            deposit = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '예수금')

            self.use_money = int(deposit) * self.use_money_percent
            # self.use_money = self.use_money / 4

            max_order_amount = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '주문가능금액')
            print('예수금: %s\n주문가능금액: %s' %
                  (int(deposit), int(max_order_amount)))
            self.detail_account_info_event_loop.exit()

        if sRQName == '계좌평가잔고내역요청':  # TR 목록의 opw00018에 있는 output 항목을 가져올 수 있음.
            total_buy_amount = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총매입금액')
            total_buy_amount_result = int(total_buy_amount)

            total_profit_amount = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총평가손익금액')
            total_profit_amount_result = int(total_profit_amount)

            total_profit_rate = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총수익률(%)')
            total_profit_rate_result = float(total_profit_rate)

            print('총매입금액: %s\n총평가손익금액: %s\n총수익률(%%): %s' % (
                total_buy_amount_result, total_profit_amount_result, total_profit_rate_result))

            # rows = self.dynamicCall(
            #     'GetRepeatCnt(QString, QString', sTrCode, sRQName) #한page에 20개까지만 불러오기 가능.
            # cnt = 0
            # for i in range(rows):
            #     code = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목번호')
            #     code_nm = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '종목명')
            #     learn_rate = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '수익률(%)')
            #     buy_price = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '매입가')
            #     current_price = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '현재가')
            #     stock_quantity = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '보유수량')
            #     possible_quantity = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '매매가능수량')
            #     total_maeip_amount = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '매입금액')
            #     current_amount = self.dynamicCall(
            #         'GetCommDate(QString, QString, int, QString)', sTrCode, sRQName, i, '평가금액')

            #     code = code.strip().[1:]  # 공란을 제외하고, 두번자 글자부터 마지막까지
            #     code_nm = code_nm.strip()
            #     learn_rate = float(learn_rate.strip())
            #     buy_price = int(buy_price.strip())
            #     current_price = int(current_price.strip())
            #     stock_quantity = int(stock_quantity.strip())
            #     possible_quantity = int(possible_quantity.strip())
            #     total_maeip_amount = int(total_maeip_amount.strip())
            #     current_amount = int(current_amount.strip())

            #     if code in self.account_stock_dict:
            #         pass

            #     else:
            #         self.account_stock_dict[code] = {}

            #     asd = self.account_stock_dict[code]

            #     asd.update({'종목명': code_nm})
            #     asd.update({'수익률(%)': learn_rate})
            #     asd.update({'매입가': buy_price})
            #     asd.update({'현재가': current_price})
            #     asd.update({'보유수량': stock_quantity})
            #     asd.update({'매매가능수량': possible_quantity})
            #     asd.update({'매입금액': total_maeip_amount})
            #     asd.update({'평가금액': current_amount})

            #     print('내 계좌에 있는 종목: %s' % self.account_stock_dict)

            #     cnt += 1

            # print('내 계좌에 있는 종목수: %s' % cnt)  # % len(self.account_stock_dict)

            # if sPrevNext == '2':
            #     self.detail_account_mystock(sPrevNext='2')
            # else :

            self.detail_account_info_event_loop.exit()

        elif sRQName == '실시간미체결현황':
            rows = self.dynamicCall(
                'GetRepeatCnt(QString, QString', sTrCode, sRQName)
            cnt = 0
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

                print('미체결현황: %s' % self.not_account_stock_dict)

            self.detail_account_info_event_loop.exit()

        # stock_code 앞에 첫문자: A 장내주식, J ELW종목, Q ETN종목 etc
        # def event_slots(self)에 있는 self.OnReceiveTrData.connect(self.trdata_slot)
        # tr 요청을 받는 구역_슬롯- 개발가이드 - 조회 실시간 데이터처리-관련함수
        # -void[OnReceiveTrData()
        # [OnReceiveTrData() 이벤트]
        # void OnReceiveTrData(
        # BSTR sScrNo,       // 화면번호
        # BSTR sRQName,      // 사용자 구분명
        # BSTR sTrCode,      // TR이름
        # BSTR sRecordName,  // 레코드 이름
        # BSTR sPrevNext,    // 연속조회 유무를 판단하는 값 0: 연속(추가조회)데이터

# 8158893311
