from pywinauto import Application
import win32gui
import win32process
import time

def verify_controls():
    print("!!! Activate Logi Window !!!")
    time.sleep(3)
    
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    
    print(f"Connecting to PID: {pid}")
    app = Application(backend="win32").connect(process=pid)
    
    # Try to target the window by handle
    dlg = app.window(handle=hwnd)
    
    print("\n--- Testing Date Edits ---")
    try:
        # Assuming MDI structure, we might need to go deeper
        # But let's try searching descendants
        edits = dlg.descendants(class_name="Edit")
        print(f"Found {len(edits)} Edit controls.")
        
        for i, edit in enumerate(edits[:5]): # Check first 5
            try:
                rect = edit.rectangle()
                text = edit.window_text()
                print(f"Edit #{i}: Text='{text}', Rect={rect}")
            except:
                print(f"Edit #{i}: Error reading")
                
    except Exception as e:
        print(f"Error finding edits: {e}")

    print("\n--- Testing Search Button ---")
    search_keywords = ["조회", "검색", "갱신", "확인"]
    
    for key in search_keywords:
        try:
            btns = dlg.descendants(class_name="Button", title=key)
            if btns:
                print(f"Found '{key}' button(s): {len(btns)}")
                for btn in btns:
                    print(f" - Rect: {btn.rectangle()}")
            else:
                print(f"'{key}' button NOT found.")
        except Exception as e:
            print(f"Error searching for '{key}': {e}")

if __name__ == "__main__":
    verify_controls()
