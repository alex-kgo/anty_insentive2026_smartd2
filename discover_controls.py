from pywinauto import Application
import win32gui
import win32process
import time

def discover():
    print("!!! Activte Logi Window !!!")
    time.sleep(3)
    
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    
    print(f"Connecting to PID: {pid}")
    
    # Try Win32 backend for legacy controls
    app = Application(backend="win32").connect(process=pid)
    dlg = app.window(handle=hwnd)
    
    # 1. Inspect Edits
    print("\n[Edit Controls]")
    edits = dlg.children(class_name="Edit")
    for i, edit in enumerate(edits):
        try:
            txt = edit.window_text()
            rect = edit.rectangle()
            print(f"Edit #{i}: Text='{txt}', Rect={rect}")
        except:
            pass
            
    # 2. Inspect Buttons
    print("\n[Button Controls]")
    buttons = dlg.children(class_name="Button")
    for i, btn in enumerate(buttons):
        try:
            txt = btn.window_text()
            rect = btn.rectangle()
            print(f"Button #{i}: Text='{txt}', Rect={rect}")
        except:
            pass
            
    # 3. Search for '조회' in all children recursively
    print("\n[Searching for '조회']")
    def check_child(window):
        try:
            txt = window.window_text()
            if "조회" in txt:
                print(f"FOUND '조회': Class='{window.class_name()}', Text='{txt}', Rect={window.rectangle()}")
        except:
            pass
        
        for child in window.children():
            check_child(child)
            
    check_child(dlg)

if __name__ == "__main__":
    discover()
