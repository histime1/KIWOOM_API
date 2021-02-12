from PyQt5.QAxContainer import *  # 응용프로그램 제어용
from PyQt5.QtCore import *
from autostock.config.errorCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        print('Kiwoom Class입니다.')
        ##### event loop 모음 ############################
        self.login_event_loop = None
        ##################################################

        ##### 변수 모음 ############################
        self.account_num = None
        ##################################################

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()

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
                         '예수금상세현황요청', 'opw00001',  '0',  '2000')
        # Open API 조회 함수를 호출해서 전문을 서버로 전송합니다.
        # CommRqData( "RQName"	,  "opw00001"	,  "0"	,  "화면번호"); 위의 2000번
        # 화면번호(=스크린번호)는 grouping을 화면번호당 100개까지 사용가능
        # 화면번호는 200개까지 만들수 있음.

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
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
        if sRQName == '예수금상세현황요청':  # TR 목록의 opw0001에 있는 output 항목을 가져올 수 있음.
            deposit = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '예수금')
            max_order_amount = self.dynamicCall(
                'GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '주문가능금액')
            print('예수금: %s\n주문가능금액: %s' %
                  (int(deposit), int(max_order_amount)))


# [자동 로그인]
# KOA StudioSA 로그인 버전 처리- 관련함수 - void OnEventConnect에 있는 내용 활용
# [OnEventConnect()이벤트]
#
# OnEventConnect(
# long nErrCode   // 로그인 상태를 전달하는데 자세한 내용은 아래 상세내용 참고
# )
#
# 로그인 처리 이벤트입니다. 성공이면 인자값 nErrCode가 0이며 에러는 다음과 같은 값이 전달됩니다.
#
# nErrCode별 상세내용
# -100 사용자 정보교환 실패
# -101 서버접속 실패
# -102 버전처리 실패

# PC registry

# _DKHOpenAPIEvents
# _DKHOpenAPI
# KHOpenAPI Control
# KHOpenAPILib
# C:\OpenAPI\KHOpenAPI.ocx
# KHOpenAPI Property Page
# KHOPENAPI.KHOpenAPICtrl.1
# KHOPENAPI.KHOPenAPICtrl.1
