import win32gui
import win32process
import time

def list_child_windows():
    print("!!! Activate Logi Window !!!")
    for i in range(3, 0, -1):
        print(f"{i}...")
        time.sleep(1)
        
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    window_title = win32gui.GetWindowText(hwnd)
    
    print(f"\nTarget Window: {hex(hwnd)} ({window_title}) PID: {pid}")
    
    child_windows = []
    def callback(chwnd, _):
        title = win32gui.GetWindowText(chwnd)
        cls = win32gui.GetClassName(chwnd)
        rect = win32gui.GetWindowRect(chwnd)
        if title or cls == "Edit" or cls == "Button": # Filter relevant ones or empty titles if class matches
             child_windows.append((chwnd, title, cls, rect))
    
    win32gui.EnumChildWindows(hwnd, callback, None)
    
    print(f"Found {len(child_windows)} child windows.")
    
    print("\n--- Buttons ---")
    for cw in child_windows:
        if cw[2] == "Button":
            print(f"Handle: {hex(cw[0])}, Title: '{cw[1]}', Rect: {cw[3]}")
            
    print("\n--- Edits ---")
    for cw in child_windows:
        if cw[2] == "Edit":
            print(f"Handle: {hex(cw[0])}, Title: '{cw[1]}', Rect: {cw[3]}")
            
    print("\n--- Static with Text ---")
    for cw in child_windows:
        if cw[2] == "Static" and cw[1]:
            print(f"Handle: {hex(cw[0])}, Title: '{cw[1]}', Rect: {cw[3]}")

if __name__ == "__main__":
    list_child_windows()
