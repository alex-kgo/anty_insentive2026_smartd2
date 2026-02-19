from pywinauto import Desktop
import win32gui
import time
import uiautomation as auto

def inspect_under_mouse():
    print("!!! MOUSE INSPECTION STARTED !!!")
    print("3초 뒤에 마우스 커서 아래에 있는 컨트롤 정보를 가져옵니다.")
    print("분석하려는 입력칸(시작일) 위에 마우스를 올려두세요!")
    
    for i in range(3, 0, -1):
        print(f"{i}...")
        time.sleep(1)
            
    print("\n--- Capturing Control Info ---")
    
    try:
        # Method 1: pywinauto (UIA)
        # Using uiautomation library directly often yields better 'ElementFromPoint'
        element = auto.ControlFromCursor()
        print(f"Name: '{element.Name}'")
        print(f"ControlTypeName: '{element.ControlTypeName}'")
        print(f"AutomationId: '{element.AutomationId}'")
        print(f"ClassName: '{element.ClassName}'")
        print(f"Rect: {element.BoundingRectangle}")
        
        # Method 2: Win32 API
        point = win32gui.GetCursorPos()
        hwnd = win32gui.WindowFromPoint(point)
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        print(f"\n[Win32 Handle Info]")
        print(f"Handle: {hex(hwnd)}")
        print(f"Title: '{title}'")
        print(f"Class: '{cls}'")
        print(f"Rect: {rect}")
        
    except Exception as e:
        print(f"Error during inspection: {e}")

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    inspect_under_mouse()
