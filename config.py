import os

# --- 경로 설정 (폴더 위치) ---
# 엑셀 파일이 저장될 기본 폴더입니다.
EXPORT_BASE_DIR = r"C:\_antigravity\anty_insentive2026_smartd2\logi_exports"
# 원본 파일 저장 폴더
RAW_DIR = os.path.join(EXPORT_BASE_DIR, "raw")
# 가공된 파일 저장 폴더
PROCESSED_DIR = os.path.join(EXPORT_BASE_DIR, "processed")
# 에러 발생 시 파일 저장 폴더
ERROR_DIR = os.path.join(EXPORT_BASE_DIR, "error")
# 로그 및 스크린샷 저장 폴더
LOG_DIR = os.path.join(EXPORT_BASE_DIR, "logs")
SCREENSHOT_DIR = os.path.join(LOG_DIR, "screens")

# --- 로지 프로그램 설정 ---

# 로지 프로그램 실행 파일 위치 (일반적으로 변경할 필요 없음)
LOGI_APP_PATH = r"C:\SmartD2\Update.exe"
LOGI_WINDOW_TITLE = "" 

# --- 엑셀 데이터 처리 설정 ---
# 엑셀의 A~G열이 어떤 데이터인지 정의합니다.
EXCEL_COLUMNS = {
    "A": "code",        # 코드
    "B": "name",        # 성명
    "C": "cust_in",     # 수신 (고객)
    "D": "driver_in",   # 수신 (기사)
    "E": "cust_out",    # 발신 (고객)
    "F": "driver_out",  # 발신 (기사)
    "G": "sum_check"    # 합계 확인용
}

# --- 구글 시트 업로드 설정 ---
# 구글 시트에 올라갈 헤더 이름입니다.
SHEET_HEADERS = ["날짜", "코드", "성명", "수신 합계", "발신 합계", "총합계"]

# --- 텔레그램 알림 설정 ---
TELEGRAM_MAX_RETRIES = 3

# --- 자동화 동작 설정 ---
# 날짜를 입력하는 방식을 선택합니다.
# 'coordinates': 화면 좌표를 찍어서 클릭 (추천)
# 'auto_id': 프로그램 내부 ID를 이용 (가끔 안될 수 있음)
DATE_SETTING_METHOD = 'coordinates' 

# --- [중요] 로지 프로그램 좌표 설정 ---
# **로지 프로그램 창 내부 기준 (상대 좌표)** 좌표를 입력해주세요.
# 좌표 확인 방법: get_mouse_position.py 파일을 실행하고, **로지 창을 클릭한 뒤** '상대좌표'를 확인하세요.
LOGI_COORDINATES = {
    # 시작 시점
    "start_year": (85, 72),
    "start_month": (105, 72),
    "start_day": (122, 72),
    "start_time": (140, 72),
    
    # 종료 시점
    "end_year": (230, 72),
    "end_month": (259, 72),
    "end_day": (277,72),
    "end_time": (295, 72),
    
    # 기타 버튼
    "search_button": (655,72), # [선택] 조회(갱신) 버튼 좌표
    "grid_click": (76,125),    # [선택] 목록 화면 아무곳이나 (우클릭용)
    "auth_popup_close": (524,494) # 셀 실행 시 뜨는 인증마법사/팝업 닫기 버튼 좌표
}

# --- 수집할 날짜 범위 ---
# 수집하고 싶은 기간을 입력하세요.
TARGET_DATE_RANGE = {
    "start": "2026-02-01 00:00", # 시작일
    "end": "2026-02-02 00:00"    # 종료일
}    
