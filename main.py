import sys
import os
import datetime
import calendar
import time
import pandas as pd
from logi_automation import LogiApp
from excel_handler import ExcelHandler
from google_sheet_manager import GoogleSheetManager
from telegram_bot import TelegramBot
import config

def get_month_dates(year_month):
    """
    해당 월의 모든 날짜에 대해 (시작일시, 종료일시) 리스트를 생성합니다.
    시작: 해당일 00:00
    종료: 다음날 00:00
    """
    year, month = map(int, year_month.split('-'))
    num_days = calendar.monthrange(year, month)[1]
    
    date_ranges = []
    for day in range(1, num_days + 1):
        start_date = datetime.datetime(year, month, day, 0, 0)
        end_date = start_date + datetime.timedelta(days=1)
        date_ranges.append((start_date, end_date))
    return date_ranges

def main():
    if len(sys.argv) < 2:
        target_month = input("작업할 월을 입력하세요 (예: 2026-02): ").strip()
    else:
        target_month = sys.argv[1]

    print(f"자동화 시작: {target_month}")
    
    # 모듈 초기화
    logi = LogiApp()
    excel = ExcelHandler()
    gsheet = GoogleSheetManager()
    bot = TelegramBot()

    try:
        # 1. 로지 프로그램 연결
        logi.connect()
        # 로그인 기능은 필요시 사용 (현재는 로그인 된 상태라고 가정)
        
        # 2. 구글 시트 연결
        gsheet.authenticate()
        sheet = gsheet.get_or_create_sheet(target_month)

        dates = get_month_dates(target_month)
        
        for start_dt, end_dt in dates:
            day_str = start_dt.strftime("%Y-%m-%d")
            print(f"날짜 처리 중: {day_str}")
            
            try:
                # 3. 데이터 조회 (기간 설정 및 갱신 버튼 클릭)
                logi.set_search_period(start_dt, end_dt)
                logi.click_search_button()
                
                # 4. 엑셀 화면 열기
                # 조회 결과가 있는지 확인하는 로직이 필요할 수 있음
                # 우선 갱신 버튼 후 엑셀 열기 진행
                
                logi.open_excel_view()
                
                # 엑셀 인증 마법사 닫기 (설정되어 있다면)
                logi.close_auth_popup()
                
                # 5. 데이터 추출 (열린 엑셀 파일에서)
                if excel.connect_to_active_excel():
                    data = excel.extract_data()
                    print(f"{len(data)}건의 데이터를 추출했습니다.")
                    
                    # 날짜 필드 추가
                    for row in data:
                        row['date'] = day_str
                    
                    # 6. 구글 시트에 업로드 (Upsert)
                    if data:
                        gsheet.upsert_data(data)
                    else:
                        print("업로드할 데이터가 없습니다.")
                        
                    excel.close_workbook()
                else:
                    print("엑셀 파일을 열지 못했습니다.")
                    
            except Exception as e:
                print(f"오류 발생 ({day_str}): {e}")
                # 오류가 발생해도 다음 날짜 처리를 위해 계속 진행
                continue

        # 7. 최종 결과 내보내기 및 알림
        print("월별 처리가 완료되었습니다. 보고서를 생성합니다...")
        
        all_data = sheet.get_all_records()
        df = pd.DataFrame(all_data)
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
        csv_filename = f"logi_calls_{target_month}_{timestamp}.csv"
        csv_path = os.path.join(config.PROCESSED_DIR, csv_filename)
        
        if not os.path.exists(config.PROCESSED_DIR):
             os.makedirs(config.PROCESSED_DIR)

        df.to_csv(csv_path, index=False, encoding='utf-8-sig') # 엑셀 호환을 위해 utf-8-sig 사용
        
        summary = f"""[로지 월 취합 완료]
월: {target_month}
총 행수: {len(df)}
상태: 성공"""
        
        bot.send_message(summary)
        # 파일 전송
        if os.path.exists(csv_path):
            bot.send_document(csv_path, caption=f"{target_month} 데이터")
        
        print("모든 작업이 완료되었습니다.")

    except Exception as e:
        print(f"치명적인 오류 발생: {e}")
        bot.send_message(f"로지 자동화 중 치명적인 오류 발생: {e}")

if __name__ == "__main__":
    main()
