import win32gui

def test_find():
    hwnd = win32gui.FindWindow("SmartD2-", None)
    print(f"FindWindow('SmartD2-', None) -> {hwnd}")
    
    # Try case variants just in case
    print(f"FindWindow('smartd2-', None) -> {win32gui.FindWindow('smartd2-', None)}")
    
    # Try finding by title substring
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if "SmartD2" in cls:
                 print(f"Found via Enum: HWND={hwnd}, Class={cls}, Title={win32gui.GetWindowText(hwnd)}")
    win32gui.EnumWindows(enum_handler, None)

if __name__ == "__main__":
    test_find()
