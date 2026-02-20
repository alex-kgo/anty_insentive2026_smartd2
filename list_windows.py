from pywinauto import Desktop
import win32gui

def list_windows():
    print("--- Listing Windows using win32gui ---")
    def enum_handler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                try:
                    print(f"Handle: {hwnd}, Title: '{title}'")
                except UnicodeEncodeError:
                    print(f"Handle: {hwnd}, Title: '{title.encode('ascii', 'ignore').decode('ascii')}' (Encoding error suppressed)")
    win32gui.EnumWindows(enum_handler, None)
    
    print("\n--- Listing Windows using pywinauto (backend='uia') ---")
    try:
        windows = Desktop(backend="uia").windows()
        for w in windows:
            print(f"Title: '{w.window_text()}'")
    except Exception as e:
        print(f"pywinauto error: {e}")

if __name__ == "__main__":
    list_windows()
