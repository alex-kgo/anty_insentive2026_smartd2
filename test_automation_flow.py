from logi_automation import LogiApp
import datetime
import time

def test_flow():
    print("Testing Logi Automation Flow...")
    app = LogiApp()
    
    try:
        # 1. Connect
        app.connect()
        print("Connected.")
        
        # 2. Set Date Range (User Requested: 2026-01-01 ~ 2026-01-02)
        start = datetime.datetime(2026, 1, 1, 0, 0)
        end = datetime.datetime(2026, 1, 2, 0, 0)
        
        app.set_search_period(start, end)
        print("Date set.")
        
        # 3. Click Search
        app.click_search_button()
        print("Search clicked.")
        
        # 4. Excel View
        # Warning: This will try to open Excel. User should verify if context menu appears.
        print("Attempting to open Excel View (Right Click Grid)...")
        app.open_excel_view()
        print("Excel View triggered (Check if Excel opens).")
        
    except Exception as e:
        print(f"Test Failed: {e}")

if __name__ == "__main__":
    test_flow()
