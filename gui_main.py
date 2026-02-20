import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import queue
import datetime
import os
import calendar
import pandas as pd

# 모듈 임포트
from logi_automation import LogiApp
from excel_handler import ExcelHandler
from google_sheet_manager import GoogleSheetManager
from telegram_bot import TelegramBot
import config

class GUIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("로지 자동화 컨트롤 센터 (Logi Automation Control Center)")
        self.root.geometry("850x750")
        self.root.configure(bg="#f0f2f5")
        
        # UI 변수 초기화
        self.mode_var = tk.StringVar(value="monthly")
        self.status_var = tk.StringVar(value="대기 중 (로지 프로그램 수동 로그인 필요)")
        
        # 위젯 참조용 (setup_ui에서 할당됨)
        self.month_entry = None
        self.start_date_entry = None
        self.end_date_entry = None
        self.monthly_label = None
        self.custom_label = None
        self.start_btn = None
        self.stop_btn = None
        self.log_widget = None
        
        # 스타일 설정
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TButton", font=("Malgun Gothic", 10), padding=6)
        self.style.configure("Stop.TButton", font=("Malgun Gothic", 10), padding=6, foreground="white", background="#d93025")
        self.style.map("Stop.TButton", background=[('active', '#b31412')])
        self.style.configure("Main.TFrame", background="#f0f2f5")
        self.style.configure("Header.TLabel", font=("Malgun Gothic", 18, "bold"), background="#f0f2f5", foreground="#1a73e8")
        self.style.configure("Mode.TRadiobutton", background="#f0f2f5", font=("Malgun Gothic", 10))
        
        self.setup_ui()
        
        # 로그 처리 설정
        self.log_queue = queue.Queue()
        self.root.after(100, self.process_logs)
        
        # 표준 출력 리다이렉션
        sys.stdout = self.LoggerWriter(self.log_queue)
        sys.stderr = self.LoggerWriter(self.log_queue)
        
        self.is_running = False
        self.stop_requested = False

    class LoggerWriter:
        def __init__(self, log_queue):
            self.log_queue = log_queue
        def write(self, message):
            if message.strip():
                timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
                self.log_queue.put(f"{timestamp} {message.strip()}\n")
        def flush(self):
            pass

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, style="Main.TFrame", padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Label(main_frame, text="로지 스마트D2 자동 취합 시스템", style="Header.TLabel")
        header.pack(pady=(0, 20))

        content_frame = ttk.LabelFrame(main_frame, text=" 작업 및 기간 설정 ", padding="15")
        content_frame.pack(fill=tk.X, pady=(0, 20))

        # 모드 선택 라디오 버튼
        mode_frame = ttk.Frame(content_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        modes = [
            ("월간 (년-월)", "monthly"),
            ("기간 지정", "custom"),
            ("최근 7일", "last7"),
            ("어제", "yesterday"),
            ("오늘", "today")
        ]
        
        for text, val in modes:
            rb = ttk.Radiobutton(mode_frame, text=text, value=val, variable=self.mode_var, 
                                 command=self.on_mode_change, style="Mode.TRadiobutton")
            rb.pack(side=tk.LEFT, padx=10)

        # 입력 필드 구역
        self.input_area = ttk.Frame(content_frame)
        self.input_area.pack(fill=tk.X, pady=5)

        # 현재 날짜 기준 설정
        current_date = datetime.datetime.now()

        # 1. 월간 입력용
        self.monthly_label = ttk.Label(self.input_area, text="대상 월 (YYYY-MM):")
        self.month_entry = ttk.Entry(self.input_area, width=15)
        
        if current_date.day >= 21:
            if current_date.month == 12:
                default_val = f"{current_date.year + 1}-01"
            else:
                default_val = f"{current_date.year}-{current_date.month + 1:02d}"
        else:
            default_val = current_date.strftime("%Y-%m")
            
        self.month_prev_btn = ttk.Button(self.input_area, text="<", width=3, command=lambda: self.change_month(-1))
        self.month_next_btn = ttk.Button(self.input_area, text=">", width=3, command=lambda: self.change_month(1))
        self.month_entry.insert(0, default_val)

        # 2. 기간 지정용 (AttributeError 방지를 위해 미리 생성)
        self.custom_label = ttk.Label(self.input_area, text="시작 ~ 종료일 (YYYY-MM-DD):")
        self.start_date_entry = ttk.Entry(self.input_area, width=15)
        self.end_date_entry = ttk.Entry(self.input_area, width=15)
        
        # 기간 지정 기본값: "지난달 21일 ~ 이번달 20일"
        # 종료일: 이번달 20일
        end_dt = current_date.replace(day=20)
        # 시작일: 지난달 21일
        first_of_this_month = current_date.replace(day=1)
        last_of_prev_month = first_of_this_month - datetime.timedelta(days=1)
        start_dt = last_of_prev_month.replace(day=21)
        
        self.start_date_entry.insert(0, start_dt.strftime("%Y-%m-%d"))
        self.end_date_entry.insert(0, end_dt.strftime("%Y-%m-%d"))

        # 실행 버튼 구역
        btn_frame = ttk.Frame(content_frame)
        btn_frame.pack(pady=(15, 0))

        self.start_btn = ttk.Button(btn_frame, text="자동화 시작", command=self.start_thread, width=15)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="중지", style="Stop.TButton", command=self.request_stop, state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # 기록창 구역
        log_frame = ttk.LabelFrame(main_frame, text=" 작업 진행 기록 ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_widget = scrolledtext.ScrolledText(log_frame, height=20, font=("Consolas", 10), bg="#ffffff", fg="#333333")
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        self.log_widget.config(state=tk.DISABLED)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=(10, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.on_mode_change()

    def on_mode_change(self):
        mode = self.mode_var.get()
        
        # 모든 가변 입력 숨기기
        for widget in [self.monthly_label, self.month_entry, self.month_prev_btn, self.month_next_btn, 
                       self.custom_label, self.start_date_entry, self.end_date_entry]:
            widget.pack_forget()

        if mode == "monthly":
            self.monthly_label.pack(side=tk.LEFT, padx=5)
            self.month_prev_btn.pack(side=tk.LEFT, padx=2)
            self.month_entry.pack(side=tk.LEFT, padx=2)
            self.month_next_btn.pack(side=tk.LEFT, padx=2)
        else:
            # 기간 지정, 최근 7일, 어제, 오늘 모두 날짜 범위를 보여줌
            current_date = datetime.datetime.now()
            today = current_date.replace(hour=0, minute=0, second=0, microsecond=0)
            
            s_dt, e_dt = today, today
            
            if mode == "today":
                s_dt, e_dt = today, today
            elif mode == "yesterday":
                s_dt = today - datetime.timedelta(days=1)
                e_dt = s_dt
            elif mode == "last7":
                s_dt = today - datetime.timedelta(days=6)
                e_dt = today
            elif mode == "custom":
                # 기간 지정은 기존에 설정된 기본값(21일~20일) 유지하거나 새로 계산
                # 여기서는 on_mode_change 시점에 입력창이 비어있지 않다면 유지
                pass
            
            if mode != "custom":
                self.start_date_entry.delete(0, tk.END)
                self.start_date_entry.insert(0, s_dt.strftime("%Y-%m-%d"))
                self.end_date_entry.delete(0, tk.END)
                self.end_date_entry.insert(0, e_dt.strftime("%Y-%m-%d"))
            
            self.custom_label.pack(side=tk.LEFT, padx=5)
            self.start_date_entry.pack(side=tk.LEFT, padx=5)
            # " ~ " 라벨 참조를 위해 on_mode_change 내에서 pack 처리
            if not hasattr(self, 'tilde_label'):
                self.tilde_label = ttk.Label(self.input_area, text=" ~ ")
            self.tilde_label.pack(side=tk.LEFT)
            self.end_date_entry.pack(side=tk.LEFT, padx=5)

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        self.log_queue.put(f"{timestamp} {message}\n")

    def process_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_widget.config(state=tk.NORMAL)
                self.log_widget.insert(tk.END, msg)
                self.log_widget.see(tk.END)
                self.log_widget.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(100, self.process_logs)

    def change_month(self, offset):
        current = self.month_entry.get().strip()
        try:
            dt = datetime.datetime.strptime(current, "%Y-%m")
            # Calculate new month
            year = dt.year
            month = dt.month + offset
            
            while month > 12:
                year += 1
                month -= 12
            while month < 1:
                year -= 1
                month += 12
            
            new_val = f"{year}-{month:02d}"
            self.month_entry.delete(0, tk.END)
            self.month_entry.insert(0, new_val)
        except:
            pass

    def request_stop(self):
        if self.is_running:
            self.stop_requested = True
            self.log("!!! 중기 요청됨. 현재 날짜 작업 완료 후 안전하게 중단합니다.")
            self.stop_btn.config(state=tk.DISABLED)

    def calculate_dates(self, mode):
        today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        ranges = []

        if mode == "today":
            ranges.append((today, today + datetime.timedelta(days=1)))
        elif mode == "yesterday":
            y = today - datetime.timedelta(days=1)
            ranges.append((y, today))
        elif mode == "last7":
            for i in range(6, -1, -1):
                d = today - datetime.timedelta(days=i)
                ranges.append((d, d + datetime.timedelta(days=1)))
        elif mode == "custom":
            s = datetime.datetime.strptime(self.start_date_entry.get().strip(), "%Y-%m-%d")
            e = datetime.datetime.strptime(self.end_date_entry.get().strip(), "%Y-%m-%d")
            curr = s
            while curr <= e:
                ranges.append((curr, curr + datetime.timedelta(days=1)))
                curr += datetime.timedelta(days=1)
        elif mode == "monthly":
            m_str = self.month_entry.get().strip()
            year, month = map(int, m_str.split('-'))
            
            # 입력된 '년-월'의 20일이 종료일
            end_dt = datetime.datetime(year, month, 20, 0, 0)
            
            # 전월 21일이 시작일
            if month == 1:
                start_dt = datetime.datetime(year - 1, 12, 21, 0, 0)
            else:
                start_dt = datetime.datetime(year, month - 1, 21, 0, 0)
                
            curr = start_dt
            while curr <= end_dt:
                ranges.append((curr, curr + datetime.timedelta(days=1)))
                curr += datetime.timedelta(days=1)
        return ranges

    def start_thread(self):
        if self.is_running: return
        
        mode = self.mode_var.get()
        try:
            # 유효성 검사
            if mode == "monthly":
                m = self.month_entry.get().strip()
                if len(m) != 7: raise Exception("형식 오류 (YYYY-MM)")
            elif mode == "custom":
                datetime.datetime.strptime(self.start_date_entry.get().strip(), "%Y-%m-%d")
                datetime.datetime.strptime(self.end_date_entry.get().strip(), "%Y-%m-%d")
        except Exception as e:
            messagebox.showerror("오류", f"입력값이 올바르지 않습니다: {e}")
            return

        self.is_running = True
        self.stop_requested = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        threading.Thread(target=self.run_automation, args=(mode,), daemon=True).start()

    def run_automation(self, mode):
        try:
            now = datetime.datetime.now()
            if mode == "monthly":
                work_month = self.month_entry.get().strip()
            else:
                work_month = now.strftime("%Y-%m")

            self.log(f"--- 자동화 작업 시작 (모드: {mode}) ---")
            
            logi = LogiApp()
            excel = ExcelHandler()
            gsheet = GoogleSheetManager()
            bot = TelegramBot()

            logi.connect()
            gsheet.authenticate()

            dates = self.calculate_dates(mode)
            if not dates:
                self.log("처리할 날짜가 없습니다.")
                return

            # 시작일-종료일 문자열 생성 (시트명으로 사용)
            period_str = f"{dates[0][0].strftime('%Y%m%d')}-{dates[-1][0].strftime('%Y%m%d')}"
            sheet = gsheet.get_or_create_sheet(period_str)
            
            self.log(f"처리 예정: 총 {len(dates)}일")

            success_days = 0
            for start_dt, end_dt in dates:
                if self.stop_requested: break
                
                day_str = start_dt.strftime("%Y-%m-%d")
                self.log(f"[{day_str}] 처리 중...")
                
                try:
                    logi.set_search_period(start_dt, end_dt)
                    logi.click_search_button()
                    logi.open_excel_view()
                    logi.close_auth_popup()
                    
                    if excel.connect_to_active_excel():
                        data = excel.extract_data()
                        if data:
                            for row in data: row['date'] = day_str
                            gsheet.upsert_data(data)
                            success_days += 1
                        excel.close_workbook()
                except Exception as e:
                    self.log(f"[{day_str}] 오류: {e}")
                    continue

            self.log(f"--- 작업 완료 ({success_days}/{len(dates)}일 성공) ---")
            
            if success_days > 0:
                all_records = sheet.get_all_records()
                df = pd.DataFrame(all_records)
                ts = now.strftime("%Y%m%d_%H%M")
                
                # 시작일-종료일 문자열 생성
                period_str = f"{dates[0][0].strftime('%Y%m%d')}-{dates[-1][0].strftime('%Y%m%d')}"
                
                csv_path = os.path.join(config.PROCESSED_DIR, f"result_{period_str}_{ts}.csv")
                if not os.path.exists(config.PROCESSED_DIR): os.makedirs(config.PROCESSED_DIR)
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                
                msg = (
                    f"[{period_str} 취합 완료]\n"
                    f"성공: {success_days}/{len(dates)}일\n"
                    f"총 데이터: {len(df)}건\n"
                    f"시트: {gsheet.spreadsheet_url}"
                )
                bot.send_message(msg)
                bot.send_document(csv_path, caption=f"{period_str} 결과 파일")

        except Exception as e:
            self.log(f"[치명적 오류] {e}")
        finally:
            self.is_running = False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = GUIApp(root)
    root.mainloop()
