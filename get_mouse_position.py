import pyautogui
import time
import os
import win32gui

def main():
    print("=== 마우스 좌표 확인 프로그램 (상대 좌표 지원) ===")
    print("1. 로지 프로그램 창을 클릭하여 활성화하세요.")
    print("2. 원하는 위치에 마우스를 올리세요.")
    print("3. 'Window Relative' (상대 좌표) 값을 기록하세요.")
    print("종료하려면 Ctrl + C 키를 누르세요.")
    print("===============================")
    
    try:
        while True:
            x, y = pyautogui.position()
            
            # Get Active Window
            hwnd = win32gui.GetForegroundWindow()
            rect = win32gui.GetWindowRect(hwnd)
            window_x = rect[0]
            window_y = rect[1]
            
            rel_x = x - window_x
            rel_y = y - window_y
            
            # 줄을 지우고 다시 출력
            # Using specific formatting to ensure it clears previous text
            print(f"\r[절대좌표: {x}, {y}]  [상대좌표(Window Relative): {rel_x}, {rel_y}]          ", end="")
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\n종료되었습니다.")
    except Exception as e:
        print(f"\n오류 발생: {e}")

if __name__ == "__main__":
    main()
