from pywinauto import Application
import time
import sys
import win32gui
import win32process

def inspect_active_window():
    print("!!! 중요 !!!")
    print("5초 카운트다운 동안 로지(Smart DII) 프로그램을 클릭하여 활성화해주세요.")
    
    for i in range(5, 0, -1):
        print(f"{i}초 전...")
        time.sleep(1)
    
    print("분석 시작...")
    
    try:
        # 1. Get Foreground Window info
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        title = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        
        print(f"Captured Window - PID: {pid}, Handle: {hwnd}, Title: '{title}', Class: '{class_name}'")
        
        if pid <= 0:
            print("Invalid PID captured. Please try again.")
            return

        # 2. Try Connecting via PID (Backend: uia)
        print("\n--- Attempting connection with backend='uia' ---")
        try:
            app = Application(backend="uia").connect(process=pid)
            # Find window by handle to be sure
            dlg = app.window(handle=hwnd)
            
            dump_filename = "logi_ui_dump_uia.txt"
            print(f"Saving 'uia' controls to {dump_filename}...")
            
            original_stdout = sys.stdout
            with open(dump_filename, "w", encoding="utf-8") as f:
                sys.stdout = f
                print(f"PID: {pid}, Title: {title}, Class: {class_name}")
                dlg.print_control_identifiers(depth=10)
            sys.stdout = original_stdout
            print("Success (uia)!")
            
        except Exception as e:
            print(f"Failed with uia: {e}")

        # 3. Try Connecting via PID (Backend: win32) - for legacy apps
        print("\n--- Attempting connection with backend='win32' ---")
        try:
            app = Application(backend="win32").connect(process=pid)
            dlg = app.window(handle=hwnd)
            
            dump_filename = "logi_ui_dump_win32.txt"
            print(f"Saving 'win32' controls to {dump_filename}...")
            
            original_stdout = sys.stdout
            with open(dump_filename, "w", encoding="utf-8") as f:
                sys.stdout = f
                print(f"PID: {pid}, Title: {title}, Class: {class_name}")
                dlg.print_control_identifiers(depth=10)
            sys.stdout = original_stdout
            print("Success (win32)!")
            
        except Exception as e:
            print(f"Failed with win32: {e}")

        print("\nDone. Please tell the agent which file was created (if any).")
        
    except Exception as e:
        print(f"Fatal Error: {e}")

if __name__ == "__main__":
    inspect_active_window()
