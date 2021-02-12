def errors(err_code):

    err_dict = {0: ('OP_ERR_NONE', '정상처리'),
                -10: ('OP_ERR_FAIL', '실패'),
                -11: ('OP_ERR_NO_CON_NUM', '조건번호 없슴'),
                -12: ('OP_ERR_MISMATCH_CON_NUM', '조건번호와 조건식 불일치'),
                -13: ('OP_ERR_MIS_REQ_EXC', '조건검색 조회요청 초과'),
                -100: ('OP_ERR_LOGIN', '사용자정보교환 실패'),
                -101: ('OP_ERR_CONNECT', '서버 접속 실패'),
                -102: ('OP_ERR_VERSION', '버전처리 실패'),
                -103: ('OP_ERR_FIREWALL', '개인방화벽 실패'),
                -104: ('OP_ERR_MEMORY', '메모리 보호실패'),
                -105: ('OP_ERR_INPUT', '함수입력값 오류'),
                -106: ('OP_ERR_SOCKET_CLOSED', '통신연결 종료'),
                -107: ('OP_ERR_SECU_MODULE', '보안모듈 오류'),
                -108: ('OP_ERR_REQ_AUTH_LOGIN', '공인인증 로그인 필요'),
                -200: ('OP_ERR_OVERFLOW', '시세조회 과부하'),
                -201: ('OP_ERR_RQ_STRUCT_FAIL', '전문작성 초기화 실패'),
                -203: ('OP_ERR_NO_DATA', '데이터 없음'),
                -204: ('OP_ERR_OVER_MAX_DATA', '조회가능한 종목수 초과-MAX 100개'),
                -205: ('OP_ERR_DATA_RCV_FAIL', '데이터 수신 실패'),
                -206: ('OP_ERR_OVER_MAX_FID', '조회가능한 FID수 초과-MAX 100개'),
                -207: ('OP_ERR_REAL_CANCEL', '실시간 해제오류'),
                -209: ('OP_ERR_INQUERY_PRICE', '시세조회제한'),
                -300: ('OP_ERR_ORD_WRONG_INPUT', '입력값 오류'),
                -301: ('OP_ERR_ORD_WRONG_ACCTNO', '계좌비밀번호 없음'),
                -302: ('OP_ERR_OTHER_ACC_USE', '타인계좌 사용오류'),
                -303: ('OP_ERR_MIS_2BILL_EXC', '주문가격이 주문착오 금액기준 초과'),
                -304: ('OP_ERR_MIS_5BILL_EXC', '주문가격이 주문착오 금액기준 초과'),
                -305: ('OP_ERR_MIS_1PER_EXC', '주문수량이 총발행주수의 1% 초과오류'),
                -306: ('OP_ERR_MIS_3PER_EXC', '주문수량은 총발행주수의 3% 초과오류'),
                -307: ('OP_ERR_SEND_FAIL', '주문전송 실패'),
                -308: ('OP_ERR_ORD_OVERFLOW', '주문전송 과부하'),
                -309: ('OP_ERR_MIS_300CNT_EXC', '주문수량 300계약 초과'),
                -310: ('OP_ERR_MIS_500CNT_EXC', '주문수량 500계약 초과'),
                -311: ('OP_ERR_ORD_TRANS_LIMIT', '주문전송제한 과부하'),
                -340: ('OP_ERR_ORD_WRONG_ACCTINFO', '계좌정보 없음'),
                -500: ('OP_ERR_ORD_SYMCODE_EMPTY', '종목코드 없음')
                }

    result = err_dict[err_code]

    return result
